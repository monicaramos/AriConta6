VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESListado 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   Icon            =   "frmTESListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCompensaciones 
      Height          =   6045
      Left            =   120
      TabIndex        =   320
      Top             =   0
      Width           =   8235
      Begin VB.CheckBox chkCompensa 
         Caption         =   "Dejar sólo importe compensacion"
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
         Left            =   960
         TabIndex        =   329
         Top             =   5370
         Width           =   4005
      End
      Begin VB.Frame FrameCambioFPCompensa 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   345
         Top             =   2400
         Width           =   7785
         Begin VB.TextBox txtDescFPago 
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
            Left            =   3540
            TabIndex        =   346
            Text            =   "Text1"
            Top             =   240
            Width           =   4125
         End
         Begin VB.TextBox txtFPago 
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
            Left            =   2370
            TabIndex        =   325
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Forma de pago vto"
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
            Index           =   49
            Left            =   120
            TabIndex        =   347
            Top             =   240
            Width           =   1920
         End
         Begin VB.Image imgFP 
            Height          =   240
            Index           =   8
            Left            =   1800
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.ComboBox cboCompensaVto 
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
         Left            =   2490
         Style           =   2  'Dropdown List
         TabIndex        =   323
         Top             =   1440
         Width           =   4245
      End
      Begin VB.TextBox txtConcpto 
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
         Left            =   2490
         TabIndex        =   328
         Text            =   "Text1"
         Top             =   4320
         Width           =   645
      End
      Begin VB.TextBox txtDescConcepto 
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
         Left            =   3210
         TabIndex        =   339
         Text            =   "Text1"
         Top             =   4320
         Width           =   4575
      End
      Begin VB.CommandButton cmdContabCompensaciones 
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
         Left            =   5700
         TabIndex        =   330
         Top             =   5370
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
         Index           =   22
         Left            =   6780
         TabIndex        =   331
         Top             =   5370
         Width           =   975
      End
      Begin VB.TextBox txtDescConcepto 
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
         Index           =   0
         Left            =   3210
         TabIndex        =   337
         Text            =   "Text1"
         Top             =   3840
         Width           =   4575
      End
      Begin VB.TextBox txtConcpto 
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
         Left            =   2490
         TabIndex        =   327
         Text            =   "Text1"
         Top             =   3840
         Width           =   645
      End
      Begin VB.TextBox txtDescDiario 
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
         Index           =   0
         Left            =   3210
         TabIndex        =   335
         Text            =   "Text1"
         Top             =   3240
         Width           =   4575
      End
      Begin VB.TextBox txtDiario 
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
         Left            =   2490
         TabIndex        =   326
         Text            =   "Text1"
         Top             =   3240
         Width           =   645
      End
      Begin VB.TextBox txtCtaBanc 
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
         Left            =   2490
         TabIndex        =   324
         Text            =   "Text1"
         Top             =   2040
         Width           =   1125
      End
      Begin VB.TextBox txtDescBanc 
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
         Index           =   2
         Left            =   3690
         TabIndex        =   333
         Text            =   "Text1"
         Top             =   2040
         Width           =   4125
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
         Index           =   23
         Left            =   2490
         TabIndex        =   322
         Top             =   840
         Width           =   1125
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   0
         Left            =   480
         Top             =   5370
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Compensa sobre Vto."
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
         Index           =   47
         Left            =   240
         TabIndex        =   342
         Top             =   1440
         Width           =   2160
      End
      Begin VB.Label Label6 
         Caption         =   "Pagos"
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
         Index           =   21
         Left            =   960
         TabIndex        =   341
         Top             =   4320
         Width           =   705
      End
      Begin VB.Label Label6 
         Caption         =   "Cobros"
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
         Index           =   20
         Left            =   960
         TabIndex        =   340
         Top             =   3840
         Width           =   765
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   1
         Left            =   1920
         Picture         =   "frmTESListado.frx":000C
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conceptos"
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
         Index           =   46
         Left            =   240
         TabIndex        =   338
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmTESListado.frx":685E
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmTESListado.frx":D0B0
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
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
         Index           =   45
         Left            =   240
         TabIndex        =   336
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
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
         Index           =   44
         Left            =   240
         TabIndex        =   334
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   2
         Left            =   1920
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   23
         Left            =   1920
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha contab."
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
         Index           =   43
         Left            =   240
         TabIndex        =   332
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Contabilización compensaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   12
         Left            =   720
         TabIndex        =   321
         Top             =   240
         Width           =   5370
      End
   End
   Begin VB.Frame FramereclaMail 
      Height          =   6735
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   10755
      Begin VB.CheckBox chkExcluirConEmail 
         Caption         =   "Excluir clientes con email (carta)"
         Height          =   255
         Left            =   7560
         TabIndex        =   110
         Top             =   5520
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   83
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   82
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox chkReclamaDevueltos 
         Caption         =   "Solo devueltos"
         Height          =   255
         Left            =   6600
         TabIndex        =   105
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   94
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   95
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   96
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   3
         Left            =   4320
         TabIndex        =   97
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   4
         Left            =   7320
         TabIndex        =   100
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   5
         Left            =   6240
         TabIndex        =   99
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   6
         Left            =   5280
         TabIndex        =   98
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPagoRec 
         Caption         =   "Check3"
         Height          =   195
         Index           =   7
         Left            =   8400
         TabIndex        =   101
         Top             =   4200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkMostrarCta 
         Caption         =   "Mostrar cuenta"
         Height          =   255
         Left            =   8520
         TabIndex        =   106
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CheckBox chkInsertarReclamas 
         Caption         =   "Insertar registros reclamaciones"
         Height          =   195
         Left            =   4800
         TabIndex        =   109
         Top             =   5550
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.TextBox txtVarios 
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   112
         Text            =   "Text1"
         Top             =   6240
         Width           =   2775
      End
      Begin VB.TextBox txtVarios 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   6240
         Width           =   4215
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Enviar por e-mail"
         Height          =   255
         Left            =   3000
         TabIndex        =   108
         Top             =   5520
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkMarcarUtlRecla 
         Caption         =   "Marcar fecha ultima reclamacion"
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   5520
         Width           =   2655
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Left            =   2040
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox txtDescCarta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   139
         Text            =   "Text2"
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtCarta 
         Height          =   285
         Left            =   2880
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   90
         Text            =   "Text1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   135
         Text            =   "Text1"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   2
         Left            =   6480
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   7320
         TabIndex        =   134
         Text            =   "Text1"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   10
         Left            =   9600
         TabIndex        =   87
         Text            =   "99/99/9999"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   9
         Left            =   7680
         TabIndex        =   86
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   8
         Left            =   600
         TabIndex        =   102
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   9480
         TabIndex        =   114
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdreclama 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   8280
         TabIndex        =   113
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   5
         Left            =   6480
         TabIndex        =   89
         Top             =   1860
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   7
         Left            =   5400
         TabIndex        =   85
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   88
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   3600
         TabIndex        =   84
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   7560
         TabIndex        =   118
         Text            =   "Text5"
         Top             =   1860
         Width           =   3075
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   117
         Text            =   "Text5"
         Top             =   1920
         Width           =   3075
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   116
         Text            =   "Text1"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   2
         Left            =   6480
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   7320
         TabIndex        =   115
         Text            =   "Text1"
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   41
         Left            =   1440
         TabIndex        =   544
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   40
         Left            =   240
         TabIndex        =   543
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   10
         Left            =   9360
         Top             =   1102
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
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
         Index           =   79
         Left            =   240
         TabIndex        =   542
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
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
         Index           =   18
         Left            =   4680
         TabIndex        =   142
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Asunto"
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
         Left            =   240
         TabIndex        =   141
         Top             =   6000
         Width           =   600
      End
      Begin VB.Image imgCarta 
         Height          =   240
         Left            =   3480
         Picture         =   "frmTESListado.frx":13902
         ToolTipText     =   "Seleccionar tipo carta"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
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
         Left            =   2160
         TabIndex        =   140
         Top             =   4680
         Width           =   360
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   3
         Left            =   840
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   2
         Left            =   6120
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   5
         Left            =   6120
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   4
         Left            =   840
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   7
         Left            =   5160
         Top             =   1102
         Width           =   240
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   3
         Left            =   840
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
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
         Index           =   15
         Left            =   240
         TabIndex        =   138
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   137
         Top             =   2805
         Width           =   465
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   2
         Left            =   6120
         Picture         =   "frmTESListado.frx":1A154
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   5520
         TabIndex        =   136
         Top             =   2805
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   7
         Left            =   8880
         TabIndex        =   133
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Vencimiento"
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
         Index           =   14
         Left            =   6720
         TabIndex        =   132
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   131
         Top             =   1095
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   9
         Left            =   7440
         Top             =   1102
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Reclamación"
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
         Index           =   13
         Left            =   240
         TabIndex        =   130
         Top             =   4680
         Width           =   1620
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   8
         Left            =   240
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   129
         Top             =   1965
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   5520
         TabIndex        =   127
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   126
         Top             =   1095
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   6
         Left            =   3360
         Top             =   1102
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "R E C L A M A C I O N E S"
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
         Index           =   2
         Left            =   2640
         TabIndex        =   125
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
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
         Left            =   240
         TabIndex        =   124
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  factura"
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
         Left            =   2760
         TabIndex        =   123
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
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
         Left            =   240
         TabIndex        =   122
         Top             =   3360
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   121
         Top             =   3645
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   5520
         TabIndex        =   120
         Top             =   3645
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
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
         Left            =   2880
         TabIndex        =   119
         Top             =   4680
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   128
         Top             =   1095
         Width           =   615
      End
   End
   Begin VB.Frame FrameCompensaAbonosCliente 
      Height          =   6735
      Left            =   120
      TabIndex        =   476
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton cmdVtoDestino 
         Height          =   375
         Index           =   1
         Left            =   240
         Picture         =   "frmTESListado.frx":209A6
         Style           =   1  'Graphical
         TabIndex        =   492
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   9120
         TabIndex        =   490
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdVtoDestino 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "frmTESListado.frx":213A8
         Style           =   1  'Graphical
         TabIndex        =   488
         Top             =   5640
         Width           =   375
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   8880
         TabIndex        =   487
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   7440
         TabIndex        =   484
         Top             =   1080
         Width           =   1215
      End
      Begin MSComctlLib.ListView lwCompenCli 
         Height          =   3975
         Left            =   240
         TabIndex        =   483
         Top             =   1560
         Width           =   10575
         _ExtentX        =   18653
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº Fact"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fec. fact."
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vto"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha Vto"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Forma pago"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cobro"
            Object.Width           =   2884
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Abonos"
            Object.Width           =   2884
         EndProperty
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   1560
         TabIndex        =   481
         Text            =   "Text5"
         Top             =   1080
         Width           =   3675
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   480
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCompensar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   8520
         TabIndex        =   478
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   36
         Left            =   9720
         TabIndex        =   477
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Imprimir hco compensacion"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   493
         Top             =   6240
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Resultado"
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
         Index           =   72
         Left            =   8160
         TabIndex        =   491
         Top             =   5685
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "  Establecer vencimiento destino"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   489
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rectifca./Abono"
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
         Index           =   71
         Left            =   8880
         TabIndex        =   486
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobro"
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
         Index           =   70
         Left            =   7440
         TabIndex        =   485
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
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
         Index           =   69
         Left            =   240
         TabIndex        =   482
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   17
         Left            =   1560
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Compensación abonos cliente"
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
         Index           =   20
         Left            =   1680
         TabIndex        =   479
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.Frame FrCobrosPendientesCli 
      Height          =   7455
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   10215
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   7
         Left            =   8880
         TabIndex        =   447
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   446
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   5
         Left            =   6600
         TabIndex        =   445
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   4
         Left            =   7800
         TabIndex        =   444
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   3
         Left            =   8880
         TabIndex        =   443
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   2
         Left            =   7800
         TabIndex        =   442
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   1
         Left            =   6600
         TabIndex        =   441
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTipPago 
         Caption         =   "Check3"
         Height          =   195
         Index           =   0
         Left            =   5400
         TabIndex        =   440
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox ChkObserva 
         Caption         =   "Mostrar observaciones del vencimiento"
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   4920
         Width           =   3375
      End
      Begin VB.ComboBox cboCobro 
         Height          =   315
         Index           =   1
         ItemData        =   "frmTESListado.frx":21DAA
         Left            =   9000
         List            =   "frmTESListado.frx":21DB7
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4560
         Width           =   975
      End
      Begin VB.ComboBox cboCobro 
         Height          =   315
         Index           =   0
         ItemData        =   "frmTESListado.frx":21DCA
         Left            =   9000
         List            =   "frmTESListado.frx":21DD7
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   4190
         Width           =   975
      End
      Begin VB.CheckBox chkApaisado 
         Caption         =   "Formato apaisado"
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   31
         Top             =   6960
         Width           =   1695
      End
      Begin VB.ComboBox cmbCuentas 
         Height          =   315
         Index           =   0
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4080
         Width           =   2775
      End
      Begin VB.CheckBox chkFormaPago 
         Caption         =   "Agrupar por forma pago"
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CheckBox chkNOremesar 
         Caption         =   "Solo marcados NO remesar"
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtNumfac 
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtNumfac 
         Height          =   285
         Index           =   0
         Left            =   6720
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   1
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   0
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   20
         Left            =   3360
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   19
         Left            =   1080
         TabIndex        =   2
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkEfectosContabilizados 
         Caption         =   "Mostrar riesgo"
         Height          =   255
         Left            =   8280
         TabIndex        =   19
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CheckBox ChkAgruparSituacion 
         Caption         =   "Agrupar por situacion vencimiento"
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   6
         Left            =   2040
         TabIndex        =   188
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   152
         Text            =   "Text1"
         Top             =   6840
         Width           =   2775
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   6840
         Width           =   735
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   149
         Text            =   "Text1"
         Top             =   6480
         Width           =   2775
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   6480
         Width           =   735
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   145
         Text            =   "Text1"
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   144
         Text            =   "Text1"
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4920
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Desglosar cliente"
         Height          =   255
         Left            =   5880
         TabIndex        =   27
         Top             =   6000
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Totalizar por fecha"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7920
         TabIndex        =   28
         Top             =   5985
         Width           =   1815
      End
      Begin VB.OptionButton optLCobros 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   1
         Left            =   7920
         TabIndex        =   26
         Top             =   5640
         Width           =   1335
      End
      Begin VB.OptionButton optLCobros 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   25
         Top             =   5640
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   7320
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   1
         Left            =   6360
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   7320
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   0
         Left            =   6360
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   36
         Text            =   "Text5"
         Top             =   3240
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   3600
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   0
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCobrosPendCli 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7560
         TabIndex        =   32
         Top             =   6960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8760
         TabIndex        =   34
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   8400
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   5520
         TabIndex        =   417
         Top             =   6360
         Width           =   4335
         Begin VB.OptionButton optCuenta 
            Caption         =   "Nombre"
            ForeColor       =   &H00004000&
            Height          =   255
            Index           =   1
            Left            =   2400
            MaskColor       =   &H00404000&
            TabIndex        =   30
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton optCuenta 
            Caption         =   "Cuenta"
            ForeColor       =   &H00004000&
            Height          =   255
            Index           =   0
            Left            =   720
            MaskColor       =   &H00404000&
            TabIndex        =   29
            Top             =   180
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   5280
         Top             =   3000
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Devuelto"
         Height          =   195
         Index           =   45
         Left            =   8280
         TabIndex        =   439
         Top             =   4600
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Recibido"
         Height          =   195
         Index           =   44
         Left            =   8280
         TabIndex        =   438
         Top             =   4230
         Width           =   705
      End
      Begin VB.Image ImageSe 
         Height          =   240
         Index           =   1
         Left            =   5880
         Picture         =   "frmTESListado.frx":21DEA
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image ImageSe 
         Height          =   240
         Index           =   0
         Left            =   5880
         Picture         =   "frmTESListado.frx":2863C
         Top             =   1200
         Width           =   240
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   7800
         Top             =   5565
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   5520
         Top             =   5565
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Cuentas"
         Height          =   195
         Index           =   40
         Left            =   1200
         TabIndex        =   318
         Top             =   4155
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número factura"
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
         Index           =   37
         Left            =   6720
         TabIndex        =   281
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   15
         Left            =   5280
         TabIndex        =   280
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   14
         Left            =   5280
         TabIndex        =   279
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
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
         Index           =   36
         Left            =   6000
         TabIndex        =   278
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha vencimiento"
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
         Index           =   35
         Left            =   240
         TabIndex        =   277
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   12
         Left            =   2400
         TabIndex        =   275
         Top             =   2205
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   20
         Left            =   3120
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   19
         Left            =   840
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   1
         Left            =   840
         Top             =   6840
         Width           =   240
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   0
         Left            =   840
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   151
         Top             =   6480
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   150
         Top             =   6840
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
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
         Index           =   20
         Left            =   240
         TabIndex        =   148
         Top             =   6120
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde dpto"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   147
         Top             =   4920
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta dpto"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   146
         Top             =   5280
         Width           =   945
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   1
         Left            =   1200
         Picture         =   "frmTESListado.frx":2EE8E
         Top             =   5280
         Width           =   240
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "frmTESListado.frx":356E0
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
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
         Index           =   19
         Left            =   240
         TabIndex        =   143
         Top             =   4560
         Width           =   1245
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   1
         Left            =   6000
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   0
         Left            =   6000
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar por"
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
         Left            =   5400
         TabIndex        =   54
         Top             =   5280
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   52
         Top             =   2685
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   5280
         TabIndex        =   50
         Top             =   2325
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
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
         Left            =   5280
         TabIndex        =   49
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmTESListado.frx":3BF32
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmTESListado.frx":42784
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Left            =   240
         TabIndex        =   44
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
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
         TabIndex        =   43
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Cobros pendientes clientes"
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
         Left            =   2520
         TabIndex        =   42
         Top             =   240
         Width           =   4890
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   3120
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   840
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   41
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   3600
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   1245
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   3285
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   9600
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha cálculo"
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
         Left            =   8400
         TabIndex        =   37
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   276
         Top             =   2205
         Width           =   615
      End
   End
   Begin VB.Frame FrameTransferencias 
      Height          =   3135
      Left            =   120
      TabIndex        =   235
      Top             =   0
      Width           =   4935
      Begin VB.CheckBox chkCartaAbonos 
         Caption         =   "Carta abonos"
         Height          =   255
         Left            =   480
         TabIndex        =   525
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   237
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   16
         Left            =   3360
         TabIndex        =   239
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   15
         Left            =   1200
         TabIndex        =   238
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   236
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   3480
         TabIndex        =   241
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2280
         TabIndex        =   240
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Index           =   29
         Left            =   120
         TabIndex        =   248
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
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
         Index           =   27
         Left            =   120
         TabIndex        =   247
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   480
         TabIndex        =   246
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   30
         Left            =   2400
         TabIndex        =   245
         Top             =   1920
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   16
         Left            =   3120
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   15
         Left            =   960
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   29
         Left            =   2400
         TabIndex        =   244
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   480
         TabIndex        =   243
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Index           =   9
         Left            =   120
         TabIndex        =   242
         Top             =   240
         Width           =   4650
      End
   End
   Begin VB.Frame FrameNorma57Importar 
      Height          =   6615
      Left            =   0
      TabIndex        =   529
      Top             =   0
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3480
         TabIndex        =   540
         Text            =   "Text1"
         Top             =   6120
         Width           =   3615
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   539
         Text            =   "Text1"
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNoram57Fich 
         Height          =   375
         Left            =   9840
         Picture         =   "frmTESListado.frx":48FD6
         Style           =   1  'Graphical
         TabIndex        =   537
         ToolTipText     =   "Leer"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdContabilizarNorma57 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7920
         TabIndex        =   536
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   42
         Left            =   9240
         TabIndex        =   534
         Top             =   6000
         Width           =   975
      End
      Begin MSComctlLib.ListView lwNorma57Importar 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   531
         Top             =   960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3836
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serie"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº Fact"
            Object.Width           =   1677
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fec. fact."
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Orden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fec Cobro"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSComctlLib.ListView lwNorma57Importar 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   533
         Top             =   3600
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3836
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2116
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº Fact"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Motivo"
            Object.Width           =   7832
         EndProperty
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   5
         Left            =   1800
         Top             =   6120
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
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
         Index           =   78
         Left            =   240
         TabIndex        =   541
         Top             =   6120
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Leer fichero bancario"
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
         Left            =   7680
         TabIndex        =   538
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimientos erroneos"
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
         Index           =   77
         Left            =   240
         TabIndex        =   535
         Top             =   3360
         Width           =   1950
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimientos encontrados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   76
         Left            =   120
         TabIndex        =   532
         Top             =   720
         Width           =   2250
      End
      Begin VB.Label Label2 
         Caption         =   "Importar fichero norma 57"
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
         Index           =   23
         Left            =   120
         TabIndex        =   530
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.Frame FrameGastosTranasferencia 
      Height          =   3255
      Left            =   120
      TabIndex        =   467
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdGastosTransfer 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   471
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtVarios 
         Height          =   1005
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   469
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   470
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   35
         Left            =   3840
         TabIndex        =   472
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Transferencia"
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
         Index           =   68
         Left            =   240
         TabIndex        =   475
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "€"
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
         Index           =   67
         Left            =   1320
         TabIndex        =   474
         Top             =   2640
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
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
         Index           =   66
         Left            =   240
         TabIndex        =   473
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Gastos transferencia"
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
         Index           =   19
         Left            =   840
         TabIndex        =   468
         Top             =   360
         Width           =   3450
      End
   End
   Begin VB.Frame FrameOperAsegComunica 
      Height          =   5655
      Left            =   120
      TabIndex        =   510
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame FrameFraPendOpAseg 
         Height          =   1455
         Left            =   120
         TabIndex        =   522
         Top             =   2520
         Width           =   4815
         Begin VB.CheckBox chkVarios 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   524
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Solo asegurados"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   523
            Top             =   720
            Value           =   1  'Checked
            Width           =   1815
         End
      End
      Begin VB.Frame FrameSelEmpre1 
         Height          =   3015
         Left            =   120
         TabIndex        =   519
         Top             =   1920
         Width           =   4815
         Begin MSComctlLib.ListView ListView3 
            Height          =   2055
            Left            =   840
            TabIndex        =   520
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3625
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
            Left            =   120
            TabIndex        =   521
            Top             =   360
            Width           =   825
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   960
            Picture         =   "frmTESListado.frx":499D8
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   1320
            Picture         =   "frmTESListado.frx":49B22
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdOperAsegComunica 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   518
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   35
         Left            =   3600
         TabIndex        =   516
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   34
         Left            =   1200
         TabIndex        =   515
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   39
         Left            =   3840
         TabIndex        =   511
         Top             =   5040
         Width           =   975
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   1
         Left            =   120
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   35
         Left            =   3360
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   39
         Left            =   2880
         TabIndex        =   517
         Top             =   1605
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   75
         Left            =   240
         TabIndex        =   514
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   34
         Left            =   960
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   38
         Left            =   360
         TabIndex        =   513
         Top             =   1605
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "XX"
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
         Height          =   345
         Index           =   22
         Left            =   2310
         TabIndex        =   512
         Top             =   480
         Width           =   390
      End
   End
   Begin VB.Frame FrameRecaudaEjec 
      Height          =   3975
      Left            =   120
      TabIndex        =   494
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdRecaudaEjecutiva 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2520
         TabIndex        =   499
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   19
         Left            =   840
         TabIndex        =   498
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2040
         TabIndex        =   507
         Text            =   "Text5"
         Top             =   2760
         Width           =   2715
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   18
         Left            =   840
         TabIndex        =   497
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2040
         TabIndex        =   504
         Text            =   "Text5"
         Top             =   2400
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   33
         Left            =   3120
         TabIndex        =   496
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   32
         Left            =   840
         TabIndex        =   495
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   38
         Left            =   3720
         TabIndex        =   500
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Recaudación ejecutiva"
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
         Height          =   345
         Index           =   21
         Left            =   840
         TabIndex        =   509
         Top             =   360
         Width           =   3210
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   47
         Left            =   120
         TabIndex        =   508
         Top             =   2760
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   19
         Left            =   600
         Picture         =   "frmTESListado.frx":49C6C
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   46
         Left            =   120
         TabIndex        =   506
         Top             =   2445
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
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
         Index           =   74
         Left            =   120
         TabIndex        =   505
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   18
         Left            =   600
         Picture         =   "frmTESListado.frx":504BE
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   33
         Left            =   2880
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   37
         Left            =   2400
         TabIndex        =   503
         Top             =   1605
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   502
         Top             =   1605
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   32
         Left            =   600
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha vencimiento"
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
         Index           =   73
         Left            =   120
         TabIndex        =   501
         Top             =   1200
         Width           =   1590
      End
   End
   Begin VB.Frame FrameListaRecep 
      Height          =   4095
      Left            =   120
      TabIndex        =   365
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CheckBox chkLstTalPag 
         Caption         =   "Justificante recepción"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   375
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtNumfac 
         Height          =   285
         Index           =   3
         Left            =   3600
         TabIndex        =   370
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtNumfac 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   369
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cboListPagare 
         Height          =   315
         ItemData        =   "frmTESListado.frx":56D10
         Left            =   1920
         List            =   "frmTESListado.frx":56D1D
         Style           =   2  'Dropdown List
         TabIndex        =   373
         Top             =   3000
         Width           =   735
      End
      Begin VB.CheckBox chkLstTalPag 
         Caption         =   "Desglosar vencimientos"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   374
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton cmdListaRecpDocum 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   376
         Top             =   3600
         Width           =   975
      End
      Begin VB.CheckBox chkLstTalPag 
         Caption         =   "Talón"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   372
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkLstTalPag 
         Caption         =   "Pagare"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   371
         Top             =   2400
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   25
         Left            =   3600
         TabIndex        =   368
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   24
         Left            =   1560
         TabIndex        =   367
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   24
         Left            =   4080
         TabIndex        =   378
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ID recepción"
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
         Index           =   63
         Left            =   240
         TabIndex        =   437
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   33
         Left            =   2880
         TabIndex        =   436
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   32
         Left            =   600
         TabIndex        =   435
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "LLevados a banco"
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
         Index           =   59
         Left            =   240
         TabIndex        =   416
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Index           =   58
         Left            =   240
         TabIndex        =   415
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   25
         Left            =   480
         TabIndex        =   380
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  recepción"
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
         Index           =   52
         Left            =   240
         TabIndex        =   379
         Top             =   600
         Width           =   1410
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   25
         Left            =   3360
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   24
         Left            =   2880
         TabIndex        =   377
         Top             =   960
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   24
         Left            =   1200
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado recepción documentos"
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
         Height          =   405
         Index           =   14
         Left            =   120
         TabIndex        =   366
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.Frame FrameDividVto 
      Height          =   2415
      Left            =   120
      TabIndex        =   403
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   406
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdDivVto 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   407
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   27
         Left            =   4200
         TabIndex        =   408
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "euros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   62
         Left            =   3240
         TabIndex        =   434
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   240
         TabIndex        =   409
         Top             =   960
         Width           =   5040
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Dividir vencimiento "
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
         Index           =   16
         Left            =   240
         TabIndex        =   405
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   56
         Left            =   240
         TabIndex        =   404
         Top             =   720
         Width           =   5040
      End
   End
   Begin VB.Frame FrameReclama 
      Height          =   3615
      Left            =   120
      TabIndex        =   418
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   422
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2520
         TabIndex        =   432
         Text            =   "Text5"
         Top             =   2160
         Width           =   2715
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   15
         Left            =   1440
         TabIndex        =   421
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2520
         TabIndex        =   429
         Text            =   "Text5"
         Top             =   1800
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   29
         Left            =   3960
         TabIndex        =   420
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   28
         Left            =   1440
         TabIndex        =   419
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdReclamas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   423
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   30
         Left            =   4320
         TabIndex        =   424
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   600
         TabIndex        =   433
         Top             =   2160
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   15
         Left            =   1200
         Picture         =   "frmTESListado.frx":56D2B
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   600
         TabIndex        =   431
         Top             =   1845
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
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
         Index           =   61
         Left            =   240
         TabIndex        =   430
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   10
         Left            =   1200
         Picture         =   "frmTESListado.frx":5D57D
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   31
         Left            =   3000
         TabIndex        =   428
         Top             =   1005
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   29
         Left            =   3720
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Historico reclamaciones"
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
         Index           =   17
         Left            =   480
         TabIndex        =   427
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   30
         Left            =   480
         TabIndex        =   426
         Top             =   1005
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   28
         Left            =   1200
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha reclamacion"
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
         Index           =   60
         Left            =   240
         TabIndex        =   425
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameProgreso 
      Height          =   1935
      Left            =   3360
      TabIndex        =   45
      Top             =   2280
      Width           =   4095
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label lbl2 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblPPAL 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameListadoCaja 
      Height          =   3495
      Left            =   120
      TabIndex        =   189
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox chkCaja 
         Caption         =   "Mostrar saldos arrastrados"
         Height          =   255
         Left            =   240
         TabIndex        =   317
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   10
         Left            =   1080
         TabIndex        =   206
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   9
         Left            =   1080
         TabIndex        =   205
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2160
         TabIndex        =   211
         Text            =   "Text5"
         Top             =   2160
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   210
         Text            =   "Text5"
         Top             =   1800
         Width           =   2715
      End
      Begin VB.CommandButton cmdCaja 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   207
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   3840
         TabIndex        =   208
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   12
         Left            =   3720
         TabIndex        =   200
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   191
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   240
         TabIndex        =   213
         Top             =   1845
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   212
         Top             =   2160
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   9
         Left            =   840
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   8
         Left            =   840
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Caja"
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
         Index           =   26
         Left            =   240
         TabIndex        =   209
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   204
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   203
         Top             =   885
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   12
         Left            =   3480
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   11
         Left            =   840
         Top             =   840
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
         Index           =   25
         Left            =   240
         TabIndex        =   202
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado caja"
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
         Index           =   6
         Left            =   240
         TabIndex        =   190
         Top             =   240
         Width           =   4650
      End
   End
   Begin VB.Frame FrameCobroGenerico 
      Height          =   2295
      Left            =   120
      TabIndex        =   311
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   316
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   13
         Left            =   600
         TabIndex        =   314
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2040
         TabIndex        =   313
         Text            =   "Text5"
         Top             =   960
         Width           =   2955
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   4080
         TabIndex        =   312
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta genérica para los vencimientos "
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
         Index           =   42
         Left            =   240
         TabIndex        =   315
         Top             =   360
         Width           =   3330
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   13
         Left            =   240
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame FrameFormaPago 
      Height          =   2415
      Left            =   120
      TabIndex        =   225
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdFormaPago 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   230
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   3960
         TabIndex        =   231
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   229
         Text            =   "Text1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   228
         Text            =   "Text1"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   227
         Text            =   "Text1"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   226
         Text            =   "Text1"
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado formas de pago"
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
         Index           =   8
         Left            =   120
         TabIndex        =   234
         Top             =   240
         Width           =   4650
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   27
         Left            =   240
         TabIndex        =   233
         Top             =   885
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   240
         TabIndex        =   232
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   5
         Left            =   840
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   4
         Left            =   840
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.Frame FrameDevEfec 
      Height          =   2535
      Left            =   120
      TabIndex        =   214
      Top             =   0
      Width           =   5055
      Begin VB.OptionButton optImpago 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   218
         Top             =   2040
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optImpago 
         Caption         =   "Fecha devolucion"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   217
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   14
         Left            =   3720
         TabIndex        =   216
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   13
         Left            =   1080
         TabIndex        =   215
         Top             =   990
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   3840
         TabIndex        =   220
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdEfecDev 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   219
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado efectos devueltos"
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
         Index           =   7
         Left            =   240
         TabIndex        =   224
         Top             =   240
         Width           =   4650
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
         Index           =   28
         Left            =   240
         TabIndex        =   223
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   14
         Left            =   3480
         Top             =   1012
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   13
         Left            =   840
         Top             =   1012
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   222
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   221
         Top             =   1005
         Width           =   615
      End
   End
   Begin VB.Frame FrameDpto 
      Height          =   3255
      Left            =   120
      TabIndex        =   163
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   3960
         TabIndex        =   172
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdDepto 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   171
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   167
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   166
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   165
         Text            =   "Text5"
         Top             =   1800
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   164
         Text            =   "Text5"
         Top             =   1440
         Width           =   2715
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado departamentos"
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
         Index           =   4
         Left            =   240
         TabIndex        =   186
         Top             =   480
         Width           =   4650
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   170
         Top             =   1485
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   240
         TabIndex        =   169
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
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
         Index           =   22
         Left            =   240
         TabIndex        =   168
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   7
         Left            =   840
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   6
         Left            =   840
         Top             =   1440
         Width           =   240
      End
   End
   Begin VB.Frame FrameAgentes 
      Height          =   2775
      Left            =   120
      TabIndex        =   153
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   162
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdAgente 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   161
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   155
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2040
         TabIndex        =   157
         Text            =   "Text1"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   156
         Text            =   "Text1"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   154
         Text            =   "Text1"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   4
         Left            =   720
         Top             =   1560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado agentes"
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
         Index           =   5
         Left            =   240
         TabIndex        =   187
         Top             =   240
         Width           =   4650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
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
         Left            =   120
         TabIndex        =   160
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   159
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   158
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image Imagente 
         Height          =   240
         Index           =   5
         Left            =   720
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame FramePrevision 
      Height          =   4935
      Left            =   120
      TabIndex        =   249
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   260
         Text            =   "Text1"
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optPrevision 
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   259
         Top             =   3240
         Width           =   1215
      End
      Begin VB.OptionButton optPrevision 
         Caption         =   "Fecha Vto"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   258
         Top             =   3240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chkPrevision 
         Caption         =   "Gastos"
         Height          =   195
         Index           =   2
         Left            =   4440
         TabIndex        =   257
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkPrevision 
         Caption         =   "Pagos"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   256
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkPrevision 
         Caption         =   "Cobros"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   255
         Top             =   2760
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   18
         Left            =   3480
         TabIndex        =   254
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   17
         Left            =   1320
         TabIndex        =   253
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   264
         Text            =   "Text1"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   252
         Text            =   "Text1"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   263
         Text            =   "Text1"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   251
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrevisionGastosCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   261
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4800
         TabIndex        =   262
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos imprevistos"
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
         Index           =   34
         Left            =   240
         TabIndex        =   274
         Top             =   3840
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblPrevInd 
         Height          =   495
         Left            =   240
         TabIndex        =   273
         Top             =   4320
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detallar"
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
         Index           =   33
         Left            =   240
         TabIndex        =   272
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar"
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
         Index           =   32
         Left            =   240
         TabIndex        =   271
         Top             =   3240
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fechas"
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
         Index           =   31
         Left            =   240
         TabIndex        =   270
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   35
         Left            =   2640
         TabIndex        =   269
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   268
         Top             =   2160
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   18
         Left            =   3240
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   17
         Left            =   1080
         Top             =   2160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   1
         Left            =   960
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   0
         Left            =   960
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
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
         Index           =   30
         Left            =   240
         TabIndex        =   267
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   266
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   32
         Left            =   240
         TabIndex        =   265
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado tesorería"
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
         Index           =   10
         Left            =   360
         TabIndex        =   250
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.Frame FrameRecepcionDocumentos 
      Height          =   4815
      Left            =   120
      TabIndex        =   348
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtCCost 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   356
         Text            =   "Text1"
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtDescCCoste 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   413
         Text            =   "Text1"
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   3120
         TabIndex        =   411
         Text            =   "Text5"
         Top             =   3480
         Width           =   3195
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   14
         Left            =   1920
         TabIndex        =   355
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkAgruparCtaPuente 
         Caption         =   "Agrupa apuntes cta puente"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   354
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton cmdRecepDocu 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   357
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   23
         Left            =   5400
         TabIndex        =   358
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtDescConcepto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2640
         TabIndex        =   363
         Text            =   "Text1"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtConcpto 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   353
         Text            =   "Text1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtConcpto 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   352
         Text            =   "Text1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDescConcepto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   360
         Text            =   "Text1"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   351
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   350
         Text            =   "Text1"
         Top             =   960
         Width           =   3735
      End
      Begin VB.Image imgCCoste 
         Height          =   240
         Index           =   0
         Left            =   1680
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Centro de coste"
         Height          =   255
         Index           =   29
         Left            =   480
         TabIndex        =   414
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cuenta"
         Height          =   255
         Index           =   28
         Left            =   480
         TabIndex        =   412
         Top             =   3480
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   14
         Left            =   1680
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
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
         Index           =   55
         Left            =   120
         TabIndex        =   410
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   3
         Left            =   1560
         Picture         =   "frmTESListado.frx":63DCF
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Haber"
         Height          =   255
         Index           =   23
         Left            =   600
         TabIndex        =   364
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "frmTESListado.frx":6A621
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conceptos"
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
         Index           =   51
         Left            =   120
         TabIndex        =   362
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "Debe"
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   361
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
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
         Index           =   50
         Left            =   120
         TabIndex        =   359
         Top             =   840
         Width           =   495
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   1
         Left            =   1560
         Picture         =   "frmTESListado.frx":70E73
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Height          =   405
         Index           =   13
         Left            =   480
         TabIndex        =   349
         Top             =   240
         Width           =   5370
      End
   End
   Begin VB.Frame FrameAseg_Bas 
      Height          =   5655
      Left            =   120
      TabIndex        =   287
      Top             =   0
      Width           =   6375
      Begin VB.Frame FrameAsegAvisos 
         Caption         =   "Avisos"
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
         Height          =   615
         Left            =   120
         TabIndex        =   466
         Top             =   4080
         Visible         =   0   'False
         Width           =   6015
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Siniestro"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   301
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Prorroga"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   300
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Falta de pago"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   299
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FrameForpa 
         Height          =   615
         Left            =   360
         TabIndex        =   400
         Top             =   4080
         Width           =   5775
         Begin VB.OptionButton optFP 
            Caption         =   "Descripción tipo pago"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   402
            Top             =   240
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton optFP 
            Caption         =   "Descripción forma pago"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   401
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame FrameASeg2 
         Height          =   855
         Left            =   1560
         TabIndex        =   397
         Top             =   3120
         Width           =   4575
         Begin VB.OptionButton optFecgaASig 
            Caption         =   "Fecha vencimiento"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   399
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optFecgaASig 
            Caption         =   "Fecha factura"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   398
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame FrOrdenAseg1 
         Height          =   855
         Left            =   120
         TabIndex        =   308
         Top             =   3120
         Width           =   5895
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   296
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   297
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "Póliza"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   298
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ordenar por"
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
            Index           =   41
            Left            =   0
            TabIndex        =   310
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdAsegBascios 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   307
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   5040
         TabIndex        =   309
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   3120
         TabIndex        =   303
         Text            =   "Text5"
         Top             =   2280
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   3120
         TabIndex        =   302
         Text            =   "Text5"
         Top             =   2640
         Width           =   2715
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   12
         Left            =   1800
         TabIndex        =   295
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   11
         Left            =   1800
         TabIndex        =   294
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   22
         Left            =   4440
         TabIndex        =   292
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   21
         Left            =   1800
         TabIndex        =   288
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   11
         Left            =   1440
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   12
         Left            =   1440
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta "
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
         Index           =   40
         Left            =   240
         TabIndex        =   306
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   39
         Left            =   840
         TabIndex        =   305
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   38
         Left            =   840
         TabIndex        =   304
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   22
         Left            =   4200
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   19
         Left            =   3600
         TabIndex        =   293
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha solicitud"
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
         Index           =   39
         Left            =   240
         TabIndex        =   291
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   21
         Left            =   1440
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   18
         Left            =   840
         TabIndex        =   290
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ccc"
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
         Index           =   11
         Left            =   240
         TabIndex        =   289
         Top             =   480
         Width           =   5970
      End
   End
   Begin VB.Frame FrameGastosFijos 
      Height          =   3615
      Left            =   2640
      TabIndex        =   449
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   31
         Left            =   5040
         TabIndex        =   457
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdGastosFijos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   456
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtGastoFijo 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   452
         Text            =   "Text1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtDescGastoFijo 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   462
         Text            =   "Text1"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtGastoFijo 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   451
         Text            =   "Text1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtDescGastoFijo 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   459
         Text            =   "Text1"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CheckBox chkDesglosaGastoFijo 
         Caption         =   "Desglosar gastos"
         Height          =   255
         Left            =   240
         TabIndex        =   455
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   31
         Left            =   4800
         TabIndex        =   454
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   30
         Left            =   2160
         TabIndex        =   453
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   35
         Left            =   3840
         TabIndex        =   465
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   34
         Left            =   1200
         TabIndex        =   464
         Top             =   2280
         Width           =   495
      End
      Begin VB.Image imgGastoFijo 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmTESListado.frx":776C5
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   17
         Left            =   720
         TabIndex        =   463
         Top             =   1440
         Width           =   495
      End
      Begin VB.Image imgGastoFijo 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmTESListado.frx":7DF17
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gasto fijo"
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
         Index           =   65
         Left            =   120
         TabIndex        =   461
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   16
         Left            =   720
         TabIndex        =   460
         Top             =   1080
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   31
         Left            =   4440
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   30
         Left            =   1800
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha cargo"
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
         Index           =   64
         Left            =   120
         TabIndex        =   458
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado gastos fijos"
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
         Index           =   18
         Left            =   720
         TabIndex        =   450
         Top             =   240
         Width           =   5370
      End
   End
   Begin VB.Frame frameListadoPagosBanco 
      Height          =   3855
      Left            =   120
      TabIndex        =   381
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkPagBanco 
         Caption         =   "Mostrar abonos"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   389
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkPagBanco 
         Caption         =   "Ordenado por fecha vencimiento"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   388
         Top             =   3000
         Width           =   3015
      End
      Begin VB.CommandButton cmdListadoPagosBanco 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   390
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2760
         TabIndex        =   394
         Text            =   "Text1"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   385
         Text            =   "Text1"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   391
         Text            =   "Text1"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtCtaBanc 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   384
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   27
         Left            =   3840
         TabIndex        =   387
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   26
         Left            =   1560
         TabIndex        =   386
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   25
         Left            =   4800
         TabIndex        =   392
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   27
         Left            =   600
         TabIndex        =   396
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   600
         TabIndex        =   395
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   4
         Left            =   1200
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   3
         Left            =   1200
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta banco"
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
         Index           =   54
         Left            =   240
         TabIndex        =   393
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha efecto"
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
         Index           =   53
         Left            =   360
         TabIndex        =   383
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   27
         Left            =   3480
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   26
         Left            =   1200
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado pagos por banco"
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
         Index           =   15
         Left            =   240
         TabIndex        =   382
         Top             =   360
         Width           =   5370
      End
   End
   Begin VB.Frame FrameListRem 
      Height          =   4935
      Left            =   120
      TabIndex        =   173
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox chkRem 
         Caption         =   "Formato banco"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   448
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CheckBox chkTipoRemesa 
         Caption         =   "Talones"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   194
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkTipoRemesa 
         Caption         =   "Pagarés"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   193
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkTipoRemesa 
         Caption         =   "Efectos"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   192
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Frame FrameOrdenRemesa 
         Height          =   975
         Left            =   360
         TabIndex        =   343
         Top             =   3000
         Width           =   4575
         Begin VB.OptionButton optOrdenRem 
            Caption         =   "Fecha vencimiento"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   196
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optOrdenRem 
            Caption         =   "Descr. cuenta "
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   198
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optOrdenRem 
            Caption         =   "Cuenta "
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   197
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optOrdenRem 
            Caption         =   "Factura"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   195
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CheckBox chkRem 
         Caption         =   "Desglosar recibos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   199
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   184
         Top             =   1995
         Width           =   975
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   183
         Top             =   1995
         Width           =   975
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   182
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtRem 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   181
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdListRem 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   201
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   3840
         TabIndex        =   185
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo remesa"
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
         Index           =   48
         Left            =   120
         TabIndex        =   344
         Top             =   2520
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   24
         Left            =   2760
         TabIndex        =   180
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   179
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   20
         Left            =   2760
         TabIndex        =   178
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   19
         Left            =   720
         TabIndex        =   177
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año remesa"
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
         Index           =   24
         Left            =   120
         TabIndex        =   176
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número remesa"
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
         Index           =   23
         Left            =   120
         TabIndex        =   175
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado remesas"
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
         Index           =   3
         Left            =   240
         TabIndex        =   174
         Top             =   240
         Width           =   4650
      End
   End
   Begin VB.Frame frpagosPendientes 
      Height          =   7215
      Left            =   120
      TabIndex        =   55
      Top             =   0
      Width           =   5415
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   360
         TabIndex        =   526
         Top             =   6120
         Width           =   4695
         Begin VB.OptionButton optMostraFP 
            Caption         =   "Forma de pago"
            Height          =   195
            Index           =   1
            Left            =   2520
            TabIndex        =   528
            Top             =   180
            Width           =   2055
         End
         Begin VB.OptionButton optMostraFP 
            Caption         =   "Tipo de pago"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   527
            Top             =   180
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.ComboBox cmbCuentas 
         Height          =   315
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   283
         Text            =   "Text1"
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox txtFPago 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtDescFPago 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   282
         Text            =   "Text1"
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CheckBox chkProv2 
         Caption         =   "Desglosar proveedor"
         Height          =   255
         Left            =   2400
         TabIndex        =   80
         Top             =   5400
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkProv 
         Caption         =   "Totalizar por fecha"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   79
         Top             =   5760
         Width           =   1935
      End
      Begin VB.OptionButton optProv 
         Caption         =   "Fecha vencimiento"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   77
         Top             =   5760
         Width           =   2175
      End
      Begin VB.OptionButton optProv 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   76
         Top             =   5400
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   1860
         TabIndex        =   63
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   65
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdPagosprov 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   64
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   59
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   58
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   2400
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   2040
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   3720
         TabIndex        =   57
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   56
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Cuentas"
         Height          =   195
         Index           =   41
         Left            =   240
         TabIndex        =   319
         Top             =   2955
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
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
         Index           =   38
         Left            =   240
         TabIndex        =   286
         Top             =   3480
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   240
         TabIndex        =   285
         Top             =   3885
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   284
         Top             =   4245
         Width           =   465
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   7
         Left            =   840
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image imgFP 
         Height          =   240
         Index           =   6
         Left            =   840
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar por"
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
         Left            =   240
         TabIndex        =   78
         Top             =   5160
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha cálculo"
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
         Left            =   240
         TabIndex        =   75
         Top             =   4680
         Width           =   1125
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   5
         Left            =   1560
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Pagos pendientes proveedores"
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
         Index           =   1
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Width           =   4890
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   73
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   72
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta proveedor"
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
         TabIndex        =   71
         Top             =   1800
         Width           =   1530
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   3
         Left            =   840
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   2
         Left            =   840
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   68
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   67
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   4
         Left            =   3420
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   840
         Top             =   1320
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
         Index           =   3
         Left            =   240
         TabIndex        =   66
         Top             =   1080
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmTESListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

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
    
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
'--monica
'Private WithEvents frmB As frmBuscaGrid
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmD As frmBasico '--monica frmDepartamentos
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmS As frmBasico '--monica frmSerie
Attribute frmS.VB_VarHelpID = -1

Dim Sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim cad As String
Dim CONT As Long
Dim I As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String


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










Private Sub cboCobro_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboCompensaVto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Check3_Click()

End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAgruparCtaPuente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub ChkAgruparSituacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkApaisado_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCaja_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub







Private Sub chkCompensa_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDesglosaGastoFijo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkEfectosContabilizados_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub chkEmail_Click()
    If chkEmail.Value = 1 Then
        Label4(17).Caption = "Asunto"
    Else
        Label4(17).Caption = "Firmante"
    End If
End Sub

Private Sub chkEmail_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkFormaPago_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkLstTalPag_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkMarcarUtlRecla_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkNOremesar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub ChkObserva_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPagBanco_Click(Index As Integer)
    Me.chkPagBanco(1).Visible = chkPagBanco(0).Value = 1  'el de abono SOLO para "tipo herbelca"
    
End Sub

Private Sub chkPagBanco_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPrevision_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkRem_Click(Index As Integer)
    Me.FrameOrdenRemesa.Visible = Me.chkRem(0).Value = 1
End Sub

Private Sub chkTipoRemesa_Click(Index As Integer)
    chkRem(1).Visible = chkTipoRemesa(0).Value = 0
End Sub



Private Sub chkTipPago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkTipPagoRec_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAgente_Click()
    Sql = "SELECT * from agentes"
    cad = ""
    RC = ""
    If txtAgente(5).Text <> "" Then
        cad = " codigo >=" & txtAgente(5).Text
        RC = "Desde " & txtAgente(5).Text & " - " & txtDescAgente(5).Text
    End If
    If txtAgente(4).Text <> "" Then
        If cad <> "" Then cad = cad & " AND "
        cad = cad & " codigo <=" & txtAgente(4).Text
        RC = RC & "      Hasta " & txtAgente(4).Text & " - " & txtDescAgente(4).Text
    End If
    
    If cad <> "" Then cad = " WHERE " & cad
    
    Sql = Sql & cad
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = "DELETE from Usuarios.zpendientes where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    Sql = "INSERT INTO Usuarios.zpendientes (codusu,  numorden,  nomforpa) VALUES (" & vUsu.Codigo & ","
    CONT = 0
    While Not Rs.EOF
        cad = Rs!Codigo & ",'" & DevNombreSQL(Rs!Nombre) & "')"
        cad = Sql & cad
        Conn.Execute cad
        CONT = CONT + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    If CONT = 0 Then
        MsgBox "Ningún dato con esos valores", vbExclamation
        Exit Sub
    End If
    
    cad = "DesdeHasta= """ & Trim(RC) & """|"
    
    With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 9
            .Show vbModal
        End With
    
    
    
End Sub

Private Sub cmdAsegBascios_Click()
Dim B As Boolean
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    Select Case Opcion
    Case 15
        B = ListAseguBasico
    Case 16
        'Listado facturacion operaciones aseguradas
        B = ListAsegFacturacion
    
    Case 17
        'Impagados
        B = ListAsegImpagos
        
    Case 18
        B = ListAsegEfectos
    Case 33
        B = AvisosAseguradora
    End Select
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If B Then
        'Impimir.
        Select Case Opcion
        Case 15
            Sql = ""
            'Cuenta
            cad = DesdeHasta("C", 11, 12)
            Sql = Trim(Sql & cad)
            
            cad = DesdeHasta("F", 21, 22, "Fec. solicitud:")
            If Sql <> "" Then cad = SaltoLinea & Trim(cad)
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            Sql = Sql & cad
            
            
            'Formulas
            cad = "Cuenta= """ & Sql & """|"
            
            'Fecha imp
            cad = cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            I = 2  'Numero parametros
            CONT = 31 'Opcion informe
        Case 16
        
            Sql = ""
            'Cuenta
            cad = DesdeHasta("C", 11, 12)
            Sql = Trim(Sql & cad)
            If Me.optFecgaASig(0).Value Then
                cad = DesdeHasta("F", 21, 22, "Fec. Fact:")
            Else
                cad = DesdeHasta("F", 21, 22, "Fec. Vto:")
            End If
            If Sql <> "" Then cad = SaltoLinea & Trim(cad)
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            Sql = Sql & cad
            
            
            'Formulas
            cad = "Cuenta= """ & Sql & """|"
            
            'Fecha imp
            cad = cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            I = 2  'Numero parametros
            CONT = 32 'Opcion informe
        
        Case 17
            Sql = ""
            'Cuenta
            cad = DesdeHasta("C", 11, 12)
            Sql = Trim(Sql & cad)
            
            cad = DesdeHasta("F", 21, 22, "Fec. Vto:")
            If Sql <> "" Then cad = SaltoLinea & Trim(cad)
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            Sql = Sql & cad
            
            
            'Formulas
            cad = "Cuenta= """ & Sql & """|"
            
            'Fecha imp
            cad = cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            I = 2  'Numero parametros
            CONT = 33 'Opcion informe
        
        Case 18
        
            Sql = ""
            'Cuenta
            cad = DesdeHasta("C", 11, 12)
            Sql = Trim(Sql & cad)
            
            cad = DesdeHasta("F", 21, 22, "Fec. Vto:")
            If Sql <> "" Then cad = SaltoLinea & Trim(cad)
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            Sql = Sql & cad
            
            
            'Formulas
            cad = "Cuenta= """ & Sql & """|"
            
            'Fecha imp
            cad = cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            I = 2  'Numero parametros
            CONT = 34 'Opcion informe
        Case 33
            Sql = ""
            'Cuenta
            cad = DesdeHasta("C", 11, 12)
            Sql = Trim(Sql & cad)
            

            cad = Trim(DesdeHasta("F", 21, 22, "Fecha aviso: "))
            If Sql <> "" Then cad = SaltoLinea & cad
            'If SQL <> "" Then cad = SaltoLinea & Trim(cad)
            Sql = Sql & cad
            
            
            'Formulas
            cad = "Cuenta= """ & Sql & """|"
            
            'Fecha imp
            cad = cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
            
            
            If Me.optAsegAvisos(0).Value Then
                Sql = "falta de pago"
            ElseIf Me.optAsegAvisos(1).Value Then
                Sql = "prorroga"
            Else
                Sql = "siniestro"
            End If
            cad = cad & "Titulo= """ & Sql & """|"
            
            I = 3  'Numero parametros
            CONT = 90 'Opcion informe
        End Select
        
        
        With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = I
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = CInt(CONT)
            .Show vbModal
        End With
    
        
    End If
End Sub

Private Sub cmdCaja_Click()

    'Listado caja
    
        'Voy a comprobar , si tiene caja y ademas si la caja es
        ' el de predeterminado O NO
        I = vUsu.Codigo Mod 100
        Sql = "predeterminado"
        cad = DevuelveDesdeBD("ctacaja", "susucaja", "codusu", CStr(I), "N", Sql)
        If cad = "" And vUsu.Nivel > 0 Then
            MsgBox "Cajas sin asignar", vbExclamation
            Exit Sub
        End If
        
        If vUsu.Nivel > 0 Then
            If Sql = "1" Then
              'CAJA PRINCIPAL, las muestra todas
              Sql = ""
            Else
              Sql = " AND slicaja.codusu = " & vUsu.Codigo Mod 100
            End If
        Else
            Sql = ""
        End If
    
        RC = CampoABD(Text3(11), "F", "feccaja", True)
        If RC <> "" Then Sql = Sql & " AND " & RC
        RC = CampoABD(Text3(12), "F", "feccaja", False)
        If RC <> "" Then Sql = Sql & " AND " & RC
         
        RC = CampoABD(txtCta(9), "T", "ctacaja", True)
        If RC <> "" Then Sql = Sql & " AND " & RC
        
        RC = CampoABD(txtCta(10), "T", "ctacaja", False)
        If RC <> "" Then Sql = Sql & " AND " & RC
        
        Sql = " AND susucaja.codusu = slicaja.codusu " & Sql
               
        Set Rs = New ADODB.Recordset
        
        cad = "Select count(*) from slicaja,susucaja where slicaja.codusu>=0 " & Sql
                            'Pongo numlinea para asi no tener k comrpobar si es AND , where o su pu... madre
        Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I = 0
        If Not Rs.EOF Then
            If DBLet(Rs.Fields(0), "N") > 0 Then I = 1
        End If
        Rs.Close
        Set Rs = Nothing
        If I = 0 Then
            MsgBox "Ningun registro con esos parametros", vbExclamation
            Exit Sub
        End If
        If chkCaja.Value = 1 Then
            I = 41  'saldo arrastrado
        Else
            I = 12  'normal
        End If
        
        If ImpirmirListadoCaja(Sql, Me.chkCaja.Value = 1) Then
            With frmImprimir
                .OtrosParametros = ""
                .NumeroParametros = 0
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = I
                .Show vbModal
            End With
        End If
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 20 Or Index = 23 Or Index >= 26 Then
        CadenaDesdeOtroForm = "" 'Por si acaso. Tiene que devolve "" para que no haga nada
    End If
    Unload Me
End Sub

Private Sub cmdCanceRemTalPag_Click()
        
    'Esta visible la cta contable. Con lo cual es OBLIGADO ponerala
    If txtCta(14).Text = "" Then
        MsgBox "Debe indicar cta " & Label4(55).Caption, vbExclamation
        Exit Sub
    End If
    If vParam.autocoste Then
        If Me.txtCCost(0).Text = "" Then
            MsgBox "Indique el centro de coste", vbExclamation
            Exit Sub
        End If
    Else
        Me.txtCCost(0).Text = ""
    End If

    CadenaDesdeOtroForm = txtCta(14).Text & "|" & Me.txtCCost(0).Text & "|"
    Unload Me
End Sub

Private Sub cmdCobrosPendCli_Click()
Dim Tot As Byte
Dim OpcionListado As Integer
    'Hago las comprobaciones
    If Text3(0).Text = "" Then
        MsgBox "Fecha cálculo no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    
    If Me.ChkAgruparSituacion.Value = 1 And Me.chkFormaPago.Value = 1 Then
        MsgBox "No puede agrupar por forma pago y por situación del vencimiento", vbExclamation
        Me.ChkAgruparSituacion.Value = 0
        Exit Sub
    End If
    
    If ChkAgruparSituacion.Value = 1 And Me.chkEfectosContabilizados.Value = 0 Then
        cad = "Los efectos remesados serán mostrados igualmente"
        MsgBox cad, vbExclamation
    
    End If
    
    
    
    'QUIEREN DETALLAR LAS CUENTAS
    CadenaDesdeOtroForm = ""
    If Me.cmbCuentas(0).ListIndex = 1 Then
        
        frmVarios.Opcion = 21
        CadenaDesdeOtroForm = Me.cmbCuentas(0).Tag
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            Me.cmbCuentas(0).ListIndex = 0
            Exit Sub
        Else
            
            Me.cmbCuentas(0).Tag = CadenaDesdeOtroForm
            GeneraComboCuentas
            Me.cmbCuentas(0).ListIndex = 2
        End If
    Else
        If Me.cmbCuentas(0).ListIndex = 2 Then CadenaDesdeOtroForm = Me.cmbCuentas(0).Tag
    End If
    
    Screen.MousePointer = vbHourglass
    If CobrosPendientesCliente(CadenaDesdeOtroForm) Then
        'Tesxto que iran
        Sql = "FECHA CALCULO: " & Text3(0).Text & "  "
        
        'Fecha fac
        cad = DesdeHasta("F", 1, 2, "F.Factura:")
        Sql = Sql & cad & " "
        
        'Fecha Vto
        cad = DesdeHasta("F", 19, 20, "F.VTO:")
        Sql = Sql & cad
        
        
        cad = ""
        If Me.cboCobro(0).ListIndex > 0 Then
            cad = cad & "["
            If Me.cboCobro(0).ListIndex > 1 Then cad = cad & "SIN "
            cad = cad & "Recibido]"
        End If
        
        If Me.cboCobro(1).ListIndex > 0 Then
            cad = cad & "["
            If Me.cboCobro(1).ListIndex > 1 Then cad = cad & "SIN "
            cad = cad & "devuelto]"
        End If
        If cad <> "" Then cad = "   " & cad
        Sql = Sql & cad
        
        
        
        'Agente
        If txtAgente(0).Text <> "" Or txtAgente(1).Text <> "" Then
            cad = "    AGENTE ("
            If txtAgente(0).Text <> "" And txtAgente(1).Text <> "" Then
                'Ha puesto los dos campos
                If txtAgente(0).Text <> txtAgente(1).Text Then
                    'SON DISTINTOS
                    cad = cad & txtAgente(0).Text & " hasta " & txtAgente(1).Text
                Else
                    cad = cad & txtAgente(0).Text & "  " & Me.txtDescAgente(0).Text
                    cad = UCase(cad)
                End If
            Else
                
                If txtAgente(0).Text <> "" Then cad = cad & " desde " & txtAgente(0).Text
                If txtAgente(1).Text <> "" Then cad = cad & " hasta " & txtAgente(1).Text
            End If
            cad = cad & ")"
            Sql = Sql & cad
        End If
            
        RC = ""
        cad = DesdeHasta("NF", 0, 1, "Nº Factura:")
        RC = RC & cad
        
        cad = DesdeHasta("S", 0, 1, "Serie:")
        RC = RC & cad
        
        
        
        
        
        
        
        
        
        If RC <> "" Then
            RC = SaltoLinea & Trim(RC)
            Sql = Sql & RC
        End If
        'Cuenta
        cad = DesdeHasta("C", 1, 0)
        If cad <> "" Then cad = SaltoLinea & Trim(cad)
        Sql = Sql & cad
        
        
        'Si lleva la cuentas seleccionadas una a una, las pondremos en el encabezado
        If Me.cmbCuentas(0).ListIndex = 2 Then
            If Me.cmbCuentas(0).Tag <> "" Then
                RC = Me.cmbCuentas(0).Tag
                cad = ""
                Do
                    I = InStr(1, RC, "|")
                    If I > 0 Then
                        If cad <> "" Then cad = cad & ","
                        cad = cad & "  " & Mid(RC, 1, I - 1)
                        RC = Mid(RC, I + 1)
                    End If
                Loop Until I = 0
                If cad <> "" Then
                    cad = SaltoLinea & "Cuentas: " & cad
                    Sql = Sql & cad
                End If
            End If
        End If
        
       
        
        'Forma pago
        cad = DesdeHasta("FP", 0, 1)
        If cad <> "" Then cad = SaltoLinea & Trim(cad)
        Sql = Sql & cad
        
        cad = PonerTipoPagoCobro_(False, False)
        Sql = Sql & cad
            
        'Si no solo NO remesar
        '---------------------
        If Me.chkNOremesar.Value = 1 Then Sql = Trim(Sql & "  SOLO marca no remesar.")
        
        'Formulas
        cad = "Cuenta= """ & Sql & """|"
        
        'Fecha imp
        cad = cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
        
        
        RC = ""
        'Totaliza
        If Me.optLCobros(0).Value Then
            Tot = Abs(Check2.Value)
        Else
            Tot = Abs(Check1.Value)
        End If
        cad = cad & "Totalizar= " & Tot & "|"
        With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            
            'Para saber cual abro
            If Me.optLCobros(0).Value Then
                If Check2.Value = 1 Then
                    '.Opcion = 1
                    OpcionListado = 1
                Else
                    '.Opcion = 3  'Sin desglosar datos cliente
                    OpcionListado = 3
                End If
            Else
                '.Opcion = 2
                OpcionListado = 2
            End If
            
            
            
            'Si agrupa por tipo de situacion
            If Me.ChkAgruparSituacion.Value = 0 And Me.chkFormaPago.Value = 0 Then
                'Si ordena por cta o nombre
                If Me.optCuenta(1).Value Then OpcionListado = OpcionListado + 70
            Else
                If Me.ChkAgruparSituacion.Value = 1 Then
                    'por cuenta o nombre
                    If Me.optCuenta(1).Value Then
                        OpcionListado = OpcionListado + 73 'del 74 al  76
                    Else
                        OpcionListado = OpcionListado + 20
                    End If
                End If
                If Me.chkFormaPago.Value = 1 Then
                    'por cuenta o nombnre
                    If Me.optCuenta(1).Value Then
                        OpcionListado = OpcionListado + 76 'del 77 al  79
                    Else
                        OpcionListado = OpcionListado + 50
                    End If
                End If
            End If


            If Me.chkApaisado(0).Value = 1 Then OpcionListado = OpcionListado + 500
    
            .Opcion = OpcionListado
             .Show vbModal
        End With

    
    End If
    Me.FrameProgreso.Visible = False
    Screen.MousePointer = vbDefault
        
    
    
End Sub

Private Function PonerTipoPagoCobro_(ParaSelect As Boolean, EsReclamacion As Boolean) As String
Dim I As Integer
Dim Sele As Integer
Dim Aux As String
Dim Visibles As Byte

    Aux = ""
    Sele = 0
    Visibles = 0
    If Not EsReclamacion Then
        For I = 0 To Me.chkTipPago.Count - 1
            If Me.chkTipPago(I).Visible Then
                Visibles = Visibles + 1
                If Me.chkTipPago(I).Value = 1 Then
                    Sele = Sele + 1
                    If ParaSelect Then
                        Aux = Aux & ", " & I
                    Else
                        Aux = Aux & "-" & Me.chkTipPago(I).Caption
                    End If
                End If
            End If
        Next
        'No ha marcado ninguno o los ha nmarcado todos. NO pongo nada
        If Sele = 0 Or Sele = Visibles Then Aux = ""
        
    Else
        '************************
        'Reclamaciones
        
        For I = 0 To Me.chkTipPagoRec.Count - 1
            If Me.chkTipPagoRec(I).Visible Then
                Visibles = Visibles + 1
                If Me.chkTipPagoRec(I).Value = 1 Then
                    Sele = Sele + 1
                    If ParaSelect Then
                        Aux = Aux & ", " & I
                    Else
                        Aux = Aux & "-" & Me.chkTipPagoRec(I).Caption
                    End If
                End If
            End If
        Next
        'No ha marcado ninguno o los ha nmarcado todos. NO pongo nada
        If Sele = 0 Or Sele = Visibles Then Aux = ""
    End If
    If Aux <> "" Then
        Aux = Mid(Aux, 2)
        Aux = "(" & Aux & ")"
    End If
    PonerTipoPagoCobro_ = Aux
End Function



Private Sub cmdCompensar_Click()
    
    cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
    If cad = "" Then
        MsgBox "No esta configurada la aplicación. Falta informe(10)", vbCritical
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
    Sql = ""
    If Len(cad) <> 1 Then
        'Ha seleccionado o cero o mas de uno
        If txtimpNoEdit(0).Text <> txtimpNoEdit(1).Text Then
            'importes distintos. Solo puede seleccionar UNO
            Sql = "Debe selecionar uno(y solo uno) como vencimiento destino"
            
        End If
    Else
        'Comprobaremos si el selecionado esta tb checked
        If Not lwCompenCli.ListItems(CONT).Checked Then
            Sql = "El vencimiento seleccionado no esta marcado"
        
        Else
            'Si el importe Cobro es mayor que abono, deberia estar
            Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
            If Importe <> 0 Then
                If Importe > 0 Then
                    'Es un abono
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) = "" Then Sql = "cobro"
                Else
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) <> "" Then Sql = "abono"
                End If
                If Sql <> "" Then Sql = "Debe marcar un " & Sql
            End If
            
        End If
    End If
    If TotalRegistros = 0 Or NumRegElim = 0 Then Sql = "Debe selecionar cobro(s) y abono(s)" & vbCrLf & Sql
        
    'Sep 2012
    'NO se pueden borrar las observaciones que ya estuvieran
    'RecuperaValor("text41csb|text42csb|text43csb|text51csb|text52csb|text53csb|text61csb|text62csb|text63csb|text71csb|text72csb|text73csb|text81csb|text82csb|text83csb|", J)
    If CONT > 0 Then
        'Hay seleccionado uno vto
        Set miRsAux = New ADODB.Recordset
        cad = "text41csb,text42csb,text43csb,text51csb,text52csb,text53csb,text61csb,text62csb,text63csb,text71csb,text72csb,text73csb,text81csb,text82csb,text83csb"
        cad = "Select " & cad & " FROM scobro WHERE numserie ='" & lwCompenCli.ListItems(CONT).Text & "' AND codfaccl="
        cad = cad & lwCompenCli.ListItems(CONT).SubItems(1) & " AND fecfaccl='" & Format(lwCompenCli.ListItems(CONT).SubItems(2), FormatoFecha)
        cad = cad & "' AND numorden = " & lwCompenCli.ListItems(CONT).SubItems(3)
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If miRsAux.EOF Then
            Sql = Sql & vbCrLf & " NO se ha encontrado el veto. destino"
        Else
            'Vamos a ver cuantos registros son
            CadenaDesdeOtroForm = ""
            RC = "0"
            For I = 0 To 14
                If DBLet(miRsAux.Fields(I), "T") = "" Then
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux.Fields(I).Name & "|"
                    RC = Val(RC) + 1
                End If
            Next I
            
                
                
            'If TotalRegistros + NumRegElim > 15 Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
            If TotalRegistros + NumRegElim > Val(RC) Then Sql = Sql & vbCrLf & "No caben los textos de los vencimientos"
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    
    
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
        Exit Sub
    Else
        If MsgBox("Seguro que desea realizar la compensación?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    
    
    
    
    Me.FrameCompensaAbonosCliente.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    RealizarCompensacionAbonosClientes
    Me.FrameCompensaAbonosCliente.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub





Private Sub cmdContabCompensaciones_Click()

    'COmprobaciones y leches
    If Me.txtConcpto(0).Text = "" Or txtDiario(0).Text = "" Or Text3(23).Text = "" Or _
        Me.txtConcpto(1).Text = "" Then
        MsgBox "Todos los campos de contabilizacion  son obligatorios", vbExclamation
        Exit Sub
    End If

    If Me.cboCompensaVto.ListIndex = 0 Then
        If Me.txtCtaBanc(2).Text = "" Then
            MsgBox "Campo banco no puede estar vacio", vbExclamation
            Exit Sub
        End If
    Else
        If Me.txtFPago(8).Text <> "" Then
            RC = DevuelveDesdeBD("codforpa", "formapago", "codforpa", txtFPago(8).Text, "N")
            If RC = "" Then
                MsgBox "No existe la forma de pago", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If FechaCorrecta2(CDate(Text3(23).Text), True) > 1 Then
        PonFoco Text3(23)
        Exit Sub
    End If

    If Me.cboCompensaVto.ListIndex = 0 Then
        'No compensa sobre ningun vencimiento.
        'No puede marcar la opcion del importe
        If chkCompensa.Value = 1 Then
            MsgBox "'Dejar sólo importe compensación' disponible cuando compense sobre un vencimiento", vbExclamation
            Exit Sub
        End If
    End If

    'Cargamos la cadena y cerramos
    CadenaDesdeOtroForm = Me.txtConcpto(0).Text & "|" & Me.txtConcpto(1).Text & "|" & txtDiario(0).Text & "|" & Text3(23).Text & "|" & Me.txtCtaBanc(2).Text & "|" & DevNombreSQL(txtDescBanc(2).Text) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.txtFPago(8).Text & "|" & Me.cboCompensaVto.ItemData(Me.cboCompensaVto.ListIndex) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.chkCompensa.Value & "|"
    Unload Me
End Sub

Private Sub cmdContabilizarNorma57_Click()
    Sql = ""
    If Me.lwNorma57Importar(0).ListItems.Count = 0 Then Sql = Sql & "-Ningun vencimiento desde el fichero" & vbCrLf
    If Me.txtCtaBanc(5).Text = "" Then Sql = Sql & "-Cuenta bancaria" & vbCrLf
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    
    'La madre de las batallas
    'El sql que mando
    Sql = "(numserie ,codfaccl,fecfaccl,numorden ) IN (select ccost,pos,nomdocum,numdiari from tmpconext "
    Sql = Sql & " WHERE codusu =" & vUsu.Codigo & " and numasien=0 ) "
    'CUIDADO. El trozo 'from tmpconext  WHERE codusu' tiene que estar extamente ASI
    '  ya que en ver cobros, si encuentro esto, pong la fecha de vencimiento la del PAGO por
    ' ventanilla que devuelve el banco y contabilizamos en funcion de esa fecha
            
            
    cad = Format(Now, "dd/mm/yyyy") & "|" & Me.txtCtaBanc(5).Text & " - " & Me.txtDescBanc(5).Text & "|0|"  'efectivo
    With frmTESVerCobrosPagos
        .ImporteGastosTarjeta_ = 0
        .OrdenacionEfectos = 3
        .vSQL = Sql
        .OrdenarEfecto = True
        .Regresar = False
        .ContabTransfer = False
        .Cobros = True
        .Tipo = 0
        .SegundoParametro = ""
        'Los textos
        .vTextos = cad
        .CodmactaUnica = ""

        .Show vbModal
    End With

    
    'Borro haya cancelado o no
    LimpiarDelProceso
End Sub

Private Sub cmdDepto_Click()

    RC = ""
    cad = ""
    If txtCta(7).Text <> "" Then
        cad = " AND departamentos.codmacta >='" & txtCta(7).Text & "'"
        RC = "Desde " & txtCta(7).Text & " - " & DtxtCta(7).Text
    End If
    
    If txtCta(8).Text <> "" Then
        cad = cad & " AND departamentos.codmacta <='" & txtCta(8).Text & "'"
        RC = RC & "  hasta " & txtCta(8).Text & " - " & DtxtCta(8).Text
    End If


    
    Sql = "select departamentos.codmacta, nommacta,dpto,descripcion from departamentos,cuentas where cuentas.codmacta=departamentos.codmacta" & cad
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql = "DELETE from Usuarios.zpendientes where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    CONT = 0
    Sql = "INSERT INTO Usuarios.zpendientes (codusu,  numorden,codforpa,  nomforpa, codmacta, nombre) VALUES (" & vUsu.Codigo & ","
    While Not Rs.EOF
        CONT = CONT + 1
        cad = CONT & "," & Rs!Dpto & ",'" & DevNombreSQL(Rs!Descripcion) & "','" & Rs!codmacta & "','" & DevNombreSQL(Rs!Nommacta) & "')"
        cad = Sql & cad
        Conn.Execute cad
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    If CONT = 0 Then
        MsgBox "Ningún dato con esos valores", vbExclamation
        Exit Sub
    End If
    
    cad = "DesdeHasta= """ & Trim(RC) & """|"
    
    With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 10
            .Show vbModal
        End With
    
End Sub

Private Sub cmdDivVto_Click()
Dim Im As Currency

    'Dividira el vto en dos. En uno dejara el importe que solicita y en el otro el resto
    'Los gastos s quedarian en uno asi como el cobrado si diera lugar
    Sql = ""
    If txtImporte(1).Text = "" Then Sql = "Ponga el importe" & vbCrLf
    
    RC = RecuperaValor(CadenaDesdeOtroForm, 3)
    Importe = CCur(RC)
    Im = ImporteFormateado(txtImporte(1).Text)
    If Im = 0 Then
        Sql = "Importe no puede ser cero"
    Else
        If Importe > 0 Then
            'Vencimiento normal
            If Im > Importe Then Sql = "Importe superior al máximo permitido(" & Importe & ")"
            
        Else
            'ABONO
            If Im > 0 Then
                Sql = "Es un abono. Importes negativos"
            Else
                If Im < Importe Then Sql = "Importe superior al máximo permitido(" & Importe & ")"
            End If
        End If
        
    End If
    
    
    If Sql = "" Then
        Set Rs = New ADODB.Recordset
        
        
        'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
        I = -1
        RC = "Select max(numorden) from scobro WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        Rs.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Rs.EOF Then
            Sql = "Error. Vencimiento NO encontrado: " & CadenaDesdeOtroForm
        Else
            I = Rs.Fields(0) + 1
        End If
        Rs.Close
        Set Rs = Nothing
        
    End If
    
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
        PonFoco txtImporte(1)
        Exit Sub
        
    Else
        Sql = "¿Desea desdoblar el vencimiento con uno de : " & Im & " euros?"
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    'OK.  a desdoblar
    Sql = "INSERT INTO scobro (`numorden`,`gastos`,impvenci,`fecultco`,`impcobro`,`recedocu`,"
    Sql = Sql & "`tiporem`,`codrem`,`anyorem`,`siturem`,reftalonpag,"
    Sql = Sql & "`numserie`,`codfaccl`,`fecfaccl`,`codmacta`,`codforpa`,`fecvenci`,`ctabanc1`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,`ctabanc2`,`text33csb`,`text41csb`,`text42csb`,`text43csb`,`text51csb`,`text52csb`,`text53csb`,`text61csb`,`text62csb`,`text63csb`,`text71csb`,`text72csb`,`text73csb`,`text81csb`,`text82csb`,`text83csb`,`ultimareclamacion`,`agente`,`departamento`,`Devuelto`,`situacionjuri`,`noremesar`,`obs`,`nomclien`,`domclien`,`pobclien`,`cpclien`,`proclien`,iban) "
    'Valores
    Sql = Sql & " SELECT " & I & ",NULL," & TransformaComasPuntos(CStr(Im)) & ",NULL,NULL,0,"
    Sql = Sql & "NULL,NULL,NULL,NULL,NULL,"
    Sql = Sql & "`numserie`,`codfaccl`,`fecfaccl`,`codmacta`,`codforpa`,`fecvenci`,`ctabanc1`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,`ctabanc2`,`text33csb`,`text41csb`,`text42csb`,`text43csb`,`text51csb`,`text52csb`,`text53csb`,`text61csb`,`text62csb`,`text63csb`,`text71csb`,`text72csb`,`text73csb`,`text81csb`,`text82csb`,"
    'text83csb`,
    Sql = Sql & "'Div vto." & Format(Now, "dd/mm/yyyy hh:nn") & "'"
    Sql = Sql & ",`ultimareclamacion`,`agente`,`departamento`,`Devuelto`,`situacionjuri`,`noremesar`,`obs`,`nomclien`,`domclien`,`pobclien`,`cpclien`,`proclien`,iban FROM "
    Sql = Sql & " scobro WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
    Sql = Sql & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
    Conn.BeginTrans
    
    'Hacemos
    CONT = 1
    If Ejecuta(Sql) Then
        'Hemos insertado. AHora updateamos el impvenci del que se queda
        If Im < 0 Then
            'Abonos
            Sql = "UPDATE scobro SET impvenci= impvenci + " & TransformaComasPuntos(CStr(Abs(Im)))
        Else
            'normal
            Sql = "UPDATE scobro SET impvenci= impvenci - " & TransformaComasPuntos(CStr(Im))
        End If
        
        Sql = Sql & " WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        Sql = Sql & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
        If Ejecuta(Sql) Then CONT = 0 'TODO BIEN ******
    End If
    'Si mal, volvemos
    If CONT = 1 Then
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        CadenaDesdeOtroForm = I
        Unload Me
    End If
    
    
End Sub

Private Sub cmdEfecDev_Click()
    'Listado de efectos devueltos
    Sql = ""
    RC = CampoABD(Text3(13), "F", "fechadev", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(Text3(14), "F", "fechadev", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    Set Rs = New ADODB.Recordset
    
    RC = "SELECT count(*) from sefecdev where numorden>=0" & Sql
    Rs.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0), "N") > 0 Then I = 1
    End If
    Rs.Close
    Set Rs = Nothing
    
    If I = 0 Then
        RC = "Ningun dato para mostrar"
        If Sql <> "" Then RC = RC & " con esos valores"
        MsgBox RC, vbExclamation
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    If ListadoEfectosDevueltos(Sql) Then
        
        cad = DesdeHasta("F", 13, 14)
        If cad <> "" Then cad = "Fecha devolución: " & cad
        cad = "Desde= """ & Trim(cad) & """|"
        If Me.optImpago(0).Value Then
            I = 13
        Else
            I = 14
        End If
        
        With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFormaPago_Click()
    cad = ""
    RC = CampoABD(txtFPago(4), "N", "codforpa", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtFPago(5), "N", "codforpa", False)
    If RC <> "" Then cad = cad & " AND " & RC
    I = 0
    If cad <> "" Then
        I = 1
        'Forma pago
        Sql = ""
        RC = DesdeHasta("FP", 4, 5)
        Sql = "Cuenta= """ & Trim(RC) & """|"
    
    Else
        I = 0
        Sql = ""
    End If
    
        
    If ListadoFormaPago(cad) Then
        With frmImprimir
            .OtrosParametros = Sql
            .NumeroParametros = I
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 27
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdGastosFijos_Click()
    If ListadoGastosFijos() Then
        With frmImprimir
            Sql = "Detalla= " & Abs(Me.chkDesglosaGastoFijo.Value) & "|DH= """ & cad & """|"
            
            
            .OtrosParametros = Sql
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 89
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdGastosTransfer_Click()

       If Me.txtImporte(2).Text = "" Then
            CadenaDesdeOtroForm = 0
       Else
            CadenaDesdeOtroForm = Me.txtImporte(2).Text
       End If
       Unload Me
End Sub

Private Sub cmdListadoPagosBanco_Click()
    If ListadoOrdenPago Then
    
        'Orden de pagos. Habran dos. El que devuelve la funcion de abajo y
        'el acabado en F que irá ordenado por fecha dentro del grupo del banco
    
        CadenaDesdeOtroForm = DevuelveDesdeBD("informe", "scryst", "codigo", 7) 'Orden de pago a bancos
        
        If Me.chkPagBanco(0).Value = 1 Then
            If CadenaDesdeOtroForm = "" Then
                MsgBox "Falta registro 7 scryst", vbExclamation
                Exit Sub
            End If
            Sql = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 4)
            Sql = Sql & "F.rpt"
            RC = App.Path & "\InformesT\" & Sql
            If Dir(RC, vbArchive) = "" Then
                MsgBox "No existe el listado ordenado por fecha. Consulte soporte técnico" & vbCrLf & "El programa continuará", vbExclamation
            Else
                CadenaDesdeOtroForm = Sql
            End If
        End If
        With frmImprimir
            .NumeroParametros = 1
            .FormulaSeleccion = "{zlistadopagos.codusu}=" & vUsu.Codigo
            
            .SoloImprimir = False
            .Opcion = 62
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdListaRecpDocum_Click()
Dim NomFile As String


    'Si marca la opcion de imprimir el justifacante de recepcion, el desglose tiene que estar marcado
    If chkLstTalPag(2).Value = 1 Then
        chkLstTalPag(1).Value = 1
        NomFile = DevuelveNombreInformeSCRYST(8, "Confir. recepcion talón")
        If NomFile = "" Then Exit Sub  'El msgbox ya lo da la funcion
        
    Else
        NomFile = ""
    End If
    
    If GeneraDatosTalPag Then
        
        RC = "FechaIMP= " & Format(Now, "dd/mm/yyyy") & "|Cuenta= "
    
        Sql = DesdeHasta("F", 24, 25, "F. Recep")
        If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
            'Solo uno seleccionado
            cad = "Talón"
            If (chkLstTalPag(0).Value = 1) Then cad = "Pagaré"
            Sql = Trim(Sql & Space(15) & "F. pago: " & cad)
        End If
        
        
        cad = DesdeHasta("NF", 2, 3, "Id. ")
        If cad <> "" Then
            Sql = Trim(Sql & Space(15) & cad)
        End If
        
        
        
        If cboListPagare.ListIndex >= 1 Then
            If cboListPagare.ListIndex = 1 Then
                cad = "Llevadas a "
            Else
                cad = "Pendientes de llevar"
            End If
            cad = cad & " banco"
            Sql = Trim(Sql & Space(15) & cad)
        End If
        Sql = RC & """" & Sql & """|"
        

        
        CadenaDesdeOtroForm = NomFile   'Por si es el ersonalizable
        With frmImprimir
            .OtrosParametros = Sql
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            If chkLstTalPag(3).Value = 1 Then
                'Si esta marcado la confirmacion recepcion
                If chkLstTalPag(2).Value = 1 Then
                    .Opcion = 87
                Else
                    .Opcion = 61
                End If
            Else
                .Opcion = 63
            End If
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub cmdListRem_Click()
Dim B As Boolean
    '-------------------------------------
    'LISTADO REMESAS
    'Utilizaremos las tablas de informes
    ' ztesoreriacomun, ztmplibrodiario
    '------------------------------------

    'Comprobaciones iniciales
    If Me.chkTipoRemesa(0).Value = 0 And Me.chkTipoRemesa(1).Value = 0 And Me.chkTipoRemesa(2).Value = 0 Then
        MsgBox "Seleccione algún tipo de remesa", vbExclamation
        Exit Sub
    End If
    
    If Me.chkTipoRemesa(0).Value = 1 And Me.chkRem(1).Value = 1 Then
        MsgBox "Formato banco para talones / pagarés", vbExclamation
        Exit Sub
    End If
    
    If chkRem(1).Value = 1 Then
        If Me.chkRem(0).Value = 1 Then MsgBox "Listado banco NO detalla vencimientos", vbExclamation
        chkRem(0).Value = 0
    End If
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If Me.chkRem(1).Value Then
        'FORMATO BANCO
        B = ListadoRemesasBanco
    Else
        'El de siempre
        B = ListadoRemesas
    End If
    If B Then
        With frmImprimir
            RC = "0"
            If Me.chkRem(1).Value = 1 Then
                RC = "1"
                I = 88
            Else
                I = 11
                If Me.chkRem(0).Value = 1 Then RC = "1"
            End If
            .OtrosParametros = "Mostrar= " & RC & "|"
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNoram57Fich_Click()

    If Me.lwNorma57Importar(0).ListItems.Count > 0 Or lwNorma57Importar(1).ListItems.Count > 0 Then
        Sql = "Ya hay un proceso . ¿ Desea importar otro archivo?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    Me.cmdContabilizarNorma57.Visible = False
    
    cd1.FileName = ""
    cd1.ShowOpen
    If cd1.FileName = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    LimpiarDelProceso
    Me.Refresh
   
    If procesarficheronorma57 Then
        
        'El fichero que ha entrado es correcto.
        'Ahora vamos a buscar los vencimientos
        If BuscarVtosNorma57 Then
            
            'AHORA cargamos los listviews
            CargaLWNorma57 True   'los correctos 'Si es que hay
            
            'Los errores
            CargaLWNorma57 False
    
    
    
            Me.cmdContabilizarNorma57.Visible = Me.lwNorma57Importar(0).ListItems.Count > 0
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOperAsegComunica_Click()
Dim B As Boolean


    'Fecha hasta
    Sql = ""
    If Text3(35).Text = "" Then Sql = Sql & "-Fecha hasta obligatoria" & vbCrLf

    If Opcion = 39 Then
            
            RC = ""
            For I = 1 To Me.ListView3.ListItems.Count
                If Me.ListView3.ListItems(I).Checked Then RC = RC & "1"
            Next
            If RC = "" Then Sql = Sql & "-Seleccione alguna empresa" & vbCrLf
            
            If Sql <> "" Then
                Sql = "Campos obligatorios: " & vbCrLf & vbCrLf & Sql
                MsgBox Sql, vbExclamation
                Exit Sub
            End If
    Else
    
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Comun para los dos
    Sql = "DELETE FROM Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Conn.Execute Sql

    
    If Opcion = 39 Then
        B = ComunicaDatosSeguro_
        I = 92
        CONT = 0
        Sql = ""
    Else
        B = GeneraDatosFrasAsegurados
        If Me.chkVarios(1).Value = 1 Then
            'Resumido
            I = 94
        Else
            I = 93
        End If
    End If
    If B Then
            Sql = ""
            CONT = 0
            RC = ""
            If Opcion <> 39 Then If Me.chkVarios(0).Value = 1 Then Sql = "SOLO asegurados"
                

            If Me.Text3(34).Text <> "" Then RC = RC & "desde " & Text3(34).Text
            If Me.Text3(35).Text <> "" Then RC = RC & "     hasta " & Text3(35).Text
            If RC <> "" Then
                RC = Trim(RC)
                RC = "Fechas : " & RC
                Sql = Trim(Sql & "       " & RC)
            End If
            
            Sql = "pDH= """ & Sql & """|"
            CONT = CONT + 1
            
            If Me.Opcion = 40 Then
                '   True: De factura ALZIRA
                '   False: vto      HERBELCA
                
                '//En el rpt DeFactura : Alzira es 1 (fra)    y herbelca es 0 (vto)
                RC = Abs(vParamT.FechaSeguroEsFra)
                Sql = Sql & "DeFactura= " & RC & "|"
                CONT = CONT + 1
            End If
    
            'Declaracion seguro
            'Cominicacion datos grupo
            If Opcion = 39 Then
                RC = ""
                For NumRegElim = 1 To Me.ListView3.ListItems.Count
                    If Me.ListView3.ListItems(NumRegElim).Checked Then
                        If RC <> "" Then RC = RC & SaltoLinea
                        RC = RC & Me.ListView3.ListItems(NumRegElim).Text
                    End If
                Next
        
                Sql = Sql & "Empresas= """ & RC & """|"
                CONT = CONT + 1
                
                
                RC = DevuelveDesdeBD("informe", "scryst", "codigo", 11) '
                If RC = "" Then
                    MsgBox "No esta configurada la aplicación. Falta informe(11)", vbCritical
                    Exit Sub
                End If
                CadenaDesdeOtroForm = RC
                
            End If
        
        
    End If
    Screen.MousePointer = vbDefault
        
    
    If B Then
        With frmImprimir
            .OtrosParametros = Sql
            .NumeroParametros = CONT
            .FormulaSeleccion = "{ztesoreriacomun.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    
    
        'rComunicaSeguro.rpt
    Else
        MsgBox "No se ha generado ningún dato", vbExclamation
    End If
    
End Sub

Private Sub cmdPagosprov_Click()
    'Hago las comprobaciones
    If Text3(5).Text = "" Then
        MsgBox "Fecha calculo no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    
    
    
   'QUIEREN DETALLAR LAS CUENTAS
    CadenaDesdeOtroForm = ""
    If Me.cmbCuentas(1).ListIndex = 1 Then
        
        frmVarios.Opcion = 21
        CadenaDesdeOtroForm = Me.cmbCuentas(1).Tag
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            Me.cmbCuentas(1).ListIndex = 0
            Exit Sub
        Else
            
            Me.cmbCuentas(1).Tag = CadenaDesdeOtroForm
            GeneraComboCuentas
            Me.cmbCuentas(1).ListIndex = 2
        End If
    Else
        If Me.cmbCuentas(1).ListIndex = 2 Then CadenaDesdeOtroForm = Me.cmbCuentas(1).Tag
    End If
    
    
    
    Screen.MousePointer = vbHourglass
    If PagosPendienteProv(CadenaDesdeOtroForm) Then
        'Tesxto que iran
        Sql = "FECHA CALCULO: " & Text3(5).Text & "  "
        
        'Fechas
        cad = DesdeHasta("F", 3, 4)
        Sql = Sql & cad
        
        'Cuenta
        cad = DesdeHasta("C", 2, 3)
        If cad <> "" Then cad = SaltoLinea & Trim(cad)
        Sql = Sql & cad
        
        
        'Si lleva la cuentas seleccionadas una a una, las pondremos en el encabezado
        If Me.cmbCuentas(1).ListIndex = 2 Then
            If Me.cmbCuentas(1).Tag <> "" Then
                RC = Me.cmbCuentas(1).Tag
                cad = ""
                Do
                    I = InStr(1, RC, "|")
                    If I > 0 Then
                        If cad <> "" Then cad = cad & ","
                        cad = cad & "  " & Mid(RC, 1, I - 1)
                        RC = Mid(RC, I + 1)
                    End If
                Loop Until I = 0
                If cad <> "" Then
                    cad = SaltoLinea & "Cuentas: " & cad
                    Sql = Sql & cad
                End If
            End If
        End If
        
        
        
        
        
        
        'Desde hasta FP
        cad = DesdeHasta("FP", 6, 7)
        If cad <> "" Then cad = SaltoLinea & Trim(cad)
        Sql = Sql & cad
        
        
        'Formulas
        cad = "Cuenta= """ & Sql & """|"
        
        'Fecha imp
        cad = cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
        
        
        
        'Totaliza
        cad = cad & "Totalizar= " & Abs(chkProv.Value) & "|"
        'marzo 2014
        cad = cad & "EsPorTipo= " & Abs(Me.optMostraFP(0).Value) & "|"
        
        
        With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = 4
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            If Me.optProv(0).Value Then
                If chkProv2.Value Then
                    .Opcion = 4
                Else
                    .Opcion = 6
                End If
            Else
                .Opcion = 5
            End If

            
            .Show vbModal
        End With

    
    End If
    Me.FrameProgreso.Visible = False
    Screen.MousePointer = vbDefault
    
    
    

End Sub



Private Sub cmdPrevisionGastosCobros_Click()


    'Borramos las lineas en usuarios
    lblPrevInd.Caption = "Preparando ..."
    lblPrevInd.Refresh
    Conn.Execute "DELETE FROM Usuarios.ztmpconext WHERE codusu =" & vUsu.Codigo
    Conn.Execute "DELETE FROM Usuarios.ztmpconextcab WHERE codusu =" & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset



    'Hacemos el selecet
    Sql = "select cuentas.codmacta,nommacta from ctabancaria,cuentas where cuentas.codmacta=ctabancaria.codmacta"
    RC = CampoABD(txtCtaBanc(0), "T", "ctabancaria.codmacta", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(txtCtaBanc(1), "T", "ctabancaria.codmacta", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TotalRegistros = 0
    While Not Rs.EOF
        '---
        If Not HacerPrevisionCuenta(Rs!codmacta, Rs!Nommacta) Then
        '---
            Sql = "DELETE FROM Usuarios.ztmpconextcab WHERE codusu =" & vUsu.Codigo
            Sql = Sql & " AND cta ='" & Rs!codmacta & "'"
            Conn.Execute Sql
        Else
            TotalRegistros = TotalRegistros + 1
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    lblPrevInd.Caption = ""
    Me.Refresh
    
    
    If TotalRegistros = 0 Then
        MsgBox "Ningun dato generado", vbExclamation
        Exit Sub
    End If
    
    If Me.optPrevision(0).Value Then
        Sql = "Fecha"
    Else
        Sql = "Tipo"
    End If
    'txtCtaBanc  txtDescBanc
    
    
    
    Sql = "Titulo= ""Informe tesorería (" & Sql & ")""|"
    'Fechas intervalor
    Sql = Sql & "Fechas= ""Fecha hasta " & Text3(18).Text & """|"
    'Cuentas
    RC = DesdeHasta("BANCO", 0, 1)
    Sql = Sql & "Cuenta= """ & RC & """|"
    Sql = Sql & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
    Sql = Sql & "NumPag= 0|"
    Sql = Sql & "Salto= 2|"

    'SQL = SQL & "MostrarAnterior= " & MostrarAnterior & "|"
    
    Screen.MousePointer = vbDefault
    With frmImprimir
        .OtrosParametros = Sql
        .NumeroParametros = 6
        .FormulaSeleccion = "{ado_lineas.codusu}=" & vUsu.Codigo
        '.SoloImprimir = True
        'Opcion dependera del combo
        .Opcion = 29
        .Show vbModal
    End With
    

    
    
    
    
End Sub

Private Function HacerPrevisionCuenta(Cta As String, Nommacta As String) As Boolean
Dim SaldoArrastrado As Currency
Dim Id As Currency
Dim IH As Currency


    HacerPrevisionCuenta = False
    
    lblPrevInd.Caption = Cta & " - " & Nommacta
    lblPrevInd.Refresh
    ' Las fechas son del periodo, luego me importa una mierda las fechas desde hasta
    '
    '
    CargaDatosConExt Cta, Now, Now, " 1 = 1", Nommacta
    
    Conn.Execute "insert into Usuarios.ztmpconextcab select * from tmpconextcab where codusu =" & vUsu.Codigo
    
    Conn.Execute "DELETE FROM tmpfaclin where codusu =" & vUsu.Codigo
    
    RC = "INSERT INTO tmpfaclin (codusu, IVA,codigo, Fecha, Cliente, cta,"
    RC = RC & " ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
    
    'PARA CADA CUENTA
    'mETEREMOS TODOS LOS REGISTROS EN LA TABLA
    '
    '           TMPFACLIN
    '
    'TANTO COBROS COMO PAGOS I GASTOS
    '
    'Luego, en funcion del orden(TIPO o fecha) los iremos insertando en la tabla, para que
    'el saldo que va arrastrando sea el correcto
    
    
       
        
    CONT = 0
    
    
    '--------------------
    'DETALLAR COBROS
    lblPrevInd.Caption = Cta & " - Cobros"
    lblPrevInd.Refresh
    Sql = " WHERE fecvenci<='" & Format(Text3(18).Text, FormatoFecha) & "'"
    Sql = Sql & " AND ctabanc1 ='" & Cta & "'"
    If chkPrevision(0).Value = 0 Then
        Sql = "select sum(impvenci),sum(impcobro),fecvenci from scobro " & Sql
        Sql = Sql & " GROUP BY fecvenci"
        
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        While Not miRsAux.EOF
        
            Id = DBLet(miRsAux.Fields(0), "N")
            IH = DBLet(miRsAux.Fields(1), "N")
            Importe = Id - IH

            If Importe <> 0 Then
                CONT = CONT + 1
                cad = "'COBROS'," & CONT & ",'" & Format(miRsAux!FecVenci, FormatoFecha) & "','COBROS PENDIENTES',NULL,"
                'HAY COBROS
                If Importe < 0 Then
                    cad = cad & "NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
                End If
                cad = RC & cad & ")"
                Conn.Execute cad
                
            End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
                
    Else
         'DETALLAR PAGOS COBROS
            '(codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
            'timporteD,timporteH, saldo
            
        'SQL = "select scobro.*,nommacta from scobro,cuentas where scobro.codmacta=cuentas.codmacta"
        'SQL = SQL & " AND fecvenci<='2006-01-01'"
         
        Sql = "select scobro.* from scobro " & Sql
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            cad = "'COBROS'," & CONT & ",'" & Format(miRsAux!FecVenci, FormatoFecha) & "','"
            'NUmero factura
            cad = cad & miRsAux!NUmSerie & miRsAux!codfaccl & "/" & miRsAux!numorden & "',"
            
            cad = cad & "'" & miRsAux!codmacta & "',"
            Importe = miRsAux!ImpVenci - DBLet(miRsAux!impcobro, "N")
            If Importe <> 0 Then
                If Importe < 0 Then
                    cad = cad & "NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
                End If
                cad = cad & ")"
                cad = RC & cad
                Conn.Execute cad
            End If
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        
    End If
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR PAGOS
    '--------------------
    '--------------------
    lblPrevInd.Caption = Cta & " - pagos"
    lblPrevInd.Refresh
    Sql = " WHERE fecefect<='" & Format(Text3(18).Text, FormatoFecha) & "'"
    Sql = Sql & " AND ctabanc1 ='" & Cta & "'"
    
    If chkPrevision(1).Value = 0 Then
        Sql = "select sum(impefect),sum(imppagad),fecefect from spagop " & Sql & " GROUP BY fecefect"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Importe = 0
        While Not miRsAux.EOF

                Id = DBLet(miRsAux.Fields(0), "N")
                IH = DBLet(miRsAux.Fields(1), "N")
                Importe = Id - IH
            
                If Importe <> 0 Then
                    CONT = CONT + 1
                    cad = "'PAGOS'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','PAGOS PENDIENTES',NULL,"
                    'HAY COBROS
                    If Importe > 0 Then
                        cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
                    Else
                        cad = cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",NULL"
                    End If
                    cad = RC & cad & ")"
                    Conn.Execute cad
                End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
    Else
         'DETALLAR PAGOS COBROS
        'codusu, IVA,codigo, Fecha, Cliente, cta,"
        ' ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
        
        'SQL = "select spagop.*,nommacta from spagop,cuentas where ctaprove=codmacta"
        'SQL = SQL & " AND fecefect<='2006-01-01'"
        
        Sql = "select spagop.* from spagop " & Sql
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            cad = "'PAGOS'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','"
            'NUmero factura
            cad = cad & DevNombreSQL(miRsAux!NumFactu) & "/" & miRsAux!numorden & "',"
            
            cad = cad & "'" & miRsAux!ctaprove & "',"
            Importe = miRsAux!ImpEfect - DBLet(miRsAux!imppagad, "N")
            If Importe <> 0 Then
                If Importe > 0 Then
                    cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
                Else
                    cad = cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",NULL"
                End If
                cad = cad & ")"
                cad = RC & cad
                Conn.Execute cad
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    
    
    
    
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR GASTOS GASTOS
    '--------------------
    '--------------------
    
    Sql = " from sgastfij,sgastfijd where sgastfij.codigo= sgastfijd.codigo"
    Sql = Sql & " and fecha >='" & Format(Now, FormatoFecha)
    Sql = Sql & "' AND fecha <='" & Format(Format(Text3(18).Text, FormatoFecha), FormatoFecha) & "'"
    Sql = Sql & " and ctaprevista='" & Cta & "'"
    
    'Desde 5 Abril 2006
    '------------------
    ' Si el gasto esta contbilizado desde la tesoreria, tiene la marca "contabilizado"
    Sql = Sql & " and contabilizado=0"
    
        ' ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
        
        'SQL = "select spagop.*,nommacta from spagop,cuentas where ctaprove=codmacta"
        'SQL = SQL & " AND fecefect<='2006-01-01'"
     
     
    'ABro el recodset aqui.
    'Si es EOF entonces no necesito abrir la pantalla, puesto
    ' que no habran gastos para seleccionar
    'Si NO es EOF entonces abro el form y entonces alli(en frmvarios)
    'recorro el recodset
    Sql = " select sgastfij.codigo,descripcion,fecha,importe " & Sql
    
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If miRsAux.EOF Then
        miRsAux.Close
    Else
        NumRegElim = CONT
        CadenaDesdeOtroForm = "Gastos cuenta: " & Nommacta & "|" & Cta & "|" & Val(chkPrevision(2).Value) & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & RC & "|"
        frmVarios.Opcion = 18
        frmVarios.Show vbModal
        Set miRsAux = New ADODB.Recordset
        CONT = NumRegElim
        Me.Refresh
    End If
    
    
    If CONT = 0 Then Exit Function
    
    lblPrevInd.Caption = Cta & " - Informe"
    lblPrevInd.Refresh
    'Cargo INFORME
    '------------------------------------------------------------------------------------------
    'Leo el  saldo inicial
    RC = "Select * from tmpconextcab where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SaldoArrastrado = 0
    If Not miRsAux.EOF Then SaldoArrastrado = DBLet(miRsAux!acumtotT, "N")
    miRsAux.Close
    
    'Si desgloso cobros, los detallo, si no hago el acumu
    RC = "INSERT INTO Usuarios.ztmpconext (codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
    RC = RC & "timporteD,timporteH, saldo) VALUES (" & vUsu.Codigo & ",'" & Cta & "','"
        
    
    
    'Ahora cogere todos los registros que estan cargados en tmpfaclin y los metere ya
    'en la tabla con los importes, ordenado como dice el option y
    'arrastrando saldo
    Sql = "select tmpfaclin.*,nommacta from tmpfaclin left join cuentas on cta=codmacta where codusu =" & vUsu.Codigo & " ORDER BY "
    'EL ORDEN
    If optPrevision(0).Value Then
        Sql = Sql & "fecha,cta"
    Else
        Sql = Sql & "cta,fecha"
    End If
    CONT = 1
    Id = 0
    IH = 0
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = Mid(miRsAux!iva, 1, 4) & "'," & CONT & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "','"
        
        
        
        If IsNull(miRsAux!Cta) Then
            'Stop
            cad = cad & "','" & DevNombreSQL(miRsAux!Cliente) & "'"
        Else
            cad = cad & Mid(DevNombreSQL(miRsAux!Cliente), 1, 10) & "',"
            If IsNull(miRsAux!Nommacta) Then
                cad = cad & "NULL"
            Else
                cad = cad & "'" & DevNombreSQL(miRsAux!Nommacta) & "'"
            End If
        End If
        If IsNull(miRsAux!Total) Then
            'VA AL DEBE
            Importe = miRsAux!ImpIva
            cad = cad & "," & TransformaComasPuntos(CStr(miRsAux!ImpIva)) & ",NULL,"
            Id = Id + Importe
        Else
            'HABER
            Importe = miRsAux!Total * -1
            cad = cad & ",NULL," & TransformaComasPuntos(CStr(miRsAux!Total)) & ","
            IH = IH + miRsAux!Total
        End If
        SaldoArrastrado = SaldoArrastrado + Importe
        cad = cad & TransformaComasPuntos(CStr(SaldoArrastrado)) & ")"
        cad = RC & cad
        Conn.Execute cad
        
        
        CONT = CONT + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ajusto los importes de la tabla tmpconextcab
    Sql = "UPDATE Usuarios.ztmpconextcab SET acumantD=acumtotD,acumantH=acumtotH,acumantT=acumtotT"
    Sql = Sql & " where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    Conn.Execute Sql
    Sql = "UPDATE Usuarios.ztmpconextcab SET acumperD=" & TransformaComasPuntos(CStr(Id))
    Sql = Sql & ", acumperH=" & TransformaComasPuntos(CStr(IH))
    Sql = Sql & ", acumperT=" & TransformaComasPuntos(CStr(Id - IH))
    Sql = Sql & ", acumtott=" & TransformaComasPuntos(CStr(SaldoArrastrado))
    
    Sql = Sql & " where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    Conn.Execute Sql
    
    HacerPrevisionCuenta = True
    
End Function

Private Sub cmdRecaudaEjecutiva_Click()
'--monica
'
'    SQL = " scobro.codmacta=cuentas.codmacta AND"
'    SQL = SQL & " fecejecutiva is null and impvenci+coalesce(gastos)-coalesce(impcobro,0)>0"
'    'Si fechvto
'    RC = CampoABD(Text3(32), "F", "fecvenci", True)
'    If RC <> "" Then SQL = SQL & " AND " & RC
'    RC = CampoABD(Text3(33), "F", "fecvenci", False)
'    If RC <> "" Then SQL = SQL & " AND " & RC
'    'Codmacta
'    RC = CampoABD(txtCta(18), "T", "scobro.codmacta", True)
'    If RC <> "" Then SQL = SQL & " AND " & RC
'    RC = CampoABD(txtCta(18), "T", "scobro.codmacta", False)
'    If RC <> "" Then SQL = SQL & " AND " & RC
'
'
'
'    'hacemos un COUNT
'    RC = DevuelveDesdeBD("count(*)", "scobro,cuentas", SQL & " AND 1", "1")
'    If RC = "" Then RC = "0"
'    If Val(RC) = 0 Then
'        MsgBox "No existen registros", vbExclamation
'        Exit Sub
'    End If
'
'    SQL = " FROM scobro,cuentas WHERE " & SQL
'
'    frmTESVarios.NumeroDocumento = SQL
'    frmTESVarios.Opcion = 29
'    frmTESVarios.Show vbModal
    
End Sub

Private Sub cmdRecepDocu_Click()
    If txtDiario(1).Text = "" Or Me.txtConcpto(2).Text = "" Or txtConcpto(3).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    If Me.Label4(55).Visible Then
        If Me.txtCta(14).Text = "" Then
            MsgBox "Cuentas " & Label4(55).Caption & " requerida", vbExclamation
            Exit Sub
        End If
        Sql = ""
        If vParam.autocoste Then
            RC = Mid(txtCta(14).Text, 1, 1)
            If RC = 6 Or RC = 7 Then
                If txtCCost(0).Text = "" Then
                    MsgBox "Centro de coste requerido", vbExclamation
                    Exit Sub
                Else
                    Sql = txtCCost(0).Text
                End If
            End If
            
                
        End If
        txtCCost(0).Text = Sql
        
    Else
        txtCCost(0).Text = ""
        Me.txtCta(14).Text = ""
    End If
    
    
    
    
    I = 0
    If Me.chkAgruparCtaPuente(0).Visible Then
        If Me.chkAgruparCtaPuente(0).Value Then I = 1
    End If
    CadenaDesdeOtroForm = txtDiario(1).Text & "|" & Me.txtConcpto(2).Text & "|" & txtConcpto(3).Text & "|" & I & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & txtCta(14).Text & "|" & txtCCost(0).Text & "|"
    
    Unload Me
End Sub

Private Sub cmdreclama_Click()
Dim NomArchivo As String
Dim Dpto As Integer
Dim EmpresaEscalona As Byte  '0 cualquiera  1 escalona
    
    
    Sql = ""
    
    'Si la fecha de reclamacion esta vacia----> mal
    If Text3(8).Text = "" Then Sql = Sql & "-Ponga la fecha de reclamación" & vbCrLf
'        MsgBox "- Ponga la fecha de reclamación", vbExclamation
'        Exit Sub
'    End If
    If txtDias.Text = "" Then Sql = Sql & "-Ponga los dias desde la ultima reclamación" & vbCrLf
'        MsgBox "Ponga los dias desde la ultima reclamación", vbExclamation
'        Exit Sub
'    End If
    
    If txtCarta.Text = "" Then Sql = Sql & "-Seleccione la carta a adjuntar" & vbCrLf
'        MsgBox "Seleccione la carta a adjuntar", vbExclamation
'        Exit Sub
'    End If
    
    
    'Si marca por email, NO puede marcar exlcuir clientes con email
    If chkEmail.Value = 1 Then
        If chkExcluirConEmail.Value = 1 Then Sql = Sql & "-En el envio de email no puede marcar la casilla 'excluir clientes con email'" & vbCrLf
'            MsgBox "En el envio de email no puede marcar la casilla 'excluir clientes con email'", vbExclamation
'            Exit Sub
'        End If
    End If
    
    If Sql <> "" Then
        Sql = "Opciones incorrectas: " & vbCrLf & vbCrLf & Sql
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    
    
    Sql = DevuelveDesdeBD("informe", "scryst", "codigo", 3) 'El tres es el tipo de docuemnto "reclamacion"

    If Sql = "" Then
            MsgBox "No existe la carta de reclamacion (3).", vbExclamation
            Exit Sub
    End If
    EmpresaEscalona = 0
    If LCase(Mid(Sql, 1, 3)) = "esc" Then EmpresaEscalona = 1
    
    NomArchivo = Sql
    Sql = App.Path & "\InformesT\" & Sql
    If Dir(Sql, vbArchive) = "" Then
        MsgBox "No se encuentra el archivo: " & Sql, vbExclamation
        Exit Sub
    End If
    
    
    
    
    'Si poner marcar como reclamacion entonces debe estar marcada la opcion
    'de insertar en las tablas de col reclamas
    If chkMarcarUtlRecla.Value = 1 Then
        If Me.chkInsertarReclamas.Value = 0 Then
            MsgBox "Debe marcar tambien la opcion de ' INSERTAR REGISTROS RECLAMACIONES '", vbExclamation
            Exit Sub
        End If
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Ahora haremos todo el proceso
    I = Val(txtDias.Text)
    I = I * -1
    Fecha = CDate(Text3(8).Text)
    Fecha = DateAdd("d", I, Fecha)
    
    'Ya tenemos en F la fecha a partir de la cual reclamamos
    'Montamos el SQL
    MontaSQLReclamacion
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not Rs.EOF
    
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    If I = 0 Then
        MsgBox "No hay valores entre las fechas", vbExclamation
        Exit Sub
    End If
    
    'No enlazamos por NIF, si no k en NIF guardaremos codmacta
    
    

    'AHora empezamos con la generacion de datos
    'Borramos el anterior
    cad = "DELETE FROm Usuarios.z347carta WHERE codusu = " & vUsu.Codigo
    Conn.Execute cad

    'Cadena insert
    cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, saludos,"
    cad = cad & "parrafo1, parrafo2, parrafo3, parrafo4, parrafo5, despedida, Asunto, contacto,Referencia) VALUES ("
    cad = cad & vUsu.Codigo
        
        
    'Monta Datos Empresa
    Rs.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    If Rs.EOF Then
        MsgBox "Error en los datos de la empresa " & vEmpresa.nomempre
        RC = ",'','','','','',''"  '6 campos
    Else
        RC = DBLet(Rs!siglasvia) & " " & DBLet(Rs!Direccion) & "  " & DBLet(Rs!numero) & ", " & DBLet(Rs!puerta)
        RC = ",'" & DBLet(Rs!nifempre) & "','" & vEmpresa.nomempre & "','" & RC & "','"
        RC = RC & DBLet(Rs!codpos) & "','" & DBLet(Rs!Poblacion) & "','" & DBLet(Rs!provincia) & "'"
    End If
    Rs.Close
    cad = cad & RC
    
    
    'Abrimos la carta
    RC = "SELECT * from scartas where codcarta = " & txtCarta.Text
    Rs.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ' saludos, parrafo1, parrafo2, parrafo3,
    'parrafo4, parrafo5, despedida, Asunto, Referencia, contacto
    RC = ""
    For I = 2 To 6
        RC = RC & ",'" & DevNombreSQL(DBLet(Rs.Fields(I))) & "'"
    Next I
    
    'Firmante , CArGO
    RC = RC & ",'" & txtVarios(0).Text & "','" & txtVarios(1).Text
    
    'Rc = Rc & "',NULL,NULL,NULL,NULL,NULL)"
    RC = RC & "',NULL,NULL,NULL)"
    cad = cad & RC
    'Cierro RS
    Rs.Close
    
    
    'Insertamos carta
    Conn.Execute cad
    
    'Para cada UNA la insertamos en la tmporal
    'Tomamos una tmp prestada
    'INSERT INTO zentrefechas (codusu, codigo, codccost, nomccost, conconam, nomconam,
    'codinmov, nominmov, fechaadq, valoradq, amortacu, fecventa, impventa, impperiodo) VALUES (
    cad = "DELETE FROM USUARIOS.zentrefechas WHERE codusu = " & vUsu.Codigo
    Conn.Execute cad
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = "INSERT INTO USUARIOS.zentrefechas(codusu,codigo,codccost,nomccost,fecventa,conconam,fechaadq"
    Sql = Sql & ",nominmov,impventa,impperiodo,valoradq,codinmov) VALUES (" & vUsu.Codigo & ","
    
    'Nuevo. Febrero 2010. Departamento ira en codinmov
    
    'Codigo
    'Clave autonumerica
    '   codccost,nomccost,fecventa,conconam
    '    numserie,codfac,fecfac,numoreden
    '  Importes
    'en fechaadq pondremos codmacta, asi luego iremos a insertar
    
    I = 1
    While Not Rs.EOF
    
        'Neuvo Febero 2010
        'Ademas de ver si me debe algo, si esta recibido NO lo puedo meter
        
        Importe = Rs!ImpVenci + DBLet(Rs!Gastos, "N") - DBLet(Rs!impcobro, "N")
        If DBLet(Rs!recedocu, "N") = 1 Then Importe = 0
        'If DBLet(Rs!recedocu, "N") = 1 And Importe > 0 Then Stop
        If Importe > 0 Then
            cad = I & ",'" & Rs!NUmSerie & "','"
            cad = cad & Rs!codfaccl & "','"
            cad = cad & Format(Rs!fecfaccl, FormatoFecha) & "',"
            cad = cad & Rs!numorden & ",'"
            cad = cad & Rs!codmacta & "','"
            'nomconam,impventa,impperiodo
            ' fec vto cobro, imp, cobrado
            cad = cad & Rs!FecVenci & "',"
            cad = cad & TransformaComasPuntos(CStr(Rs!ImpVenci)) & ","
            If IsNull(Rs!impcobro) Then
                cad = cad & "NULL"
            Else
                cad = cad & TransformaComasPuntos(CStr(Rs!impcobro))
            End If
            'ValorADQ=GASTOS
            cad = cad & "," & TransformaComasPuntos(CStr(DBLet(Rs!Gastos, "N")))
            
            'Febrero 2010
            'Departamento
            cad = cad & "," & DBLet(Rs!departamento, "N")
            cad = Sql & cad & ")"
            Conn.Execute cad
            
            I = I + 1
            
        End If
        Rs.MoveNext
        
    Wend
    Rs.Close
    
    If I = 1 Then
        'Ningun valor con esa opcion
        MsgBox "No hay valores entre las fechas", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'AHora ya tenemos en entrefechas todos los valores k vamos a reclamar.
    'Para ello haremos un par cosas:
    ' 1.- Para cada codmacta(fechaadq) haremos su entrada en 347 cargando su datos NIF,dir,...
    ' 2.- UPDATEAREMOS nomconam con el NIF, para en el informe enalzar
    ' 3.- tabla cuentas. Donde guardaremos los datos de la cuenta bancaria
    
    cad = "DELETE FROM Usuarios.z347  where codusu = " & vUsu.Codigo
    Conn.Execute cad
    
    cad = "DELETE FROM Usuarios.zcuentas  where codusu = " & vUsu.Codigo
    Conn.Execute cad
    
    cad = "SELECT fechaadq,codinmov FROM USUARIOS.zentrefechas WHERE codusu = " & vUsu.Codigo & " GROUP BY fechaadq,codinmov"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Datos contables
    Set miRsAux = New ADODB.Recordset
    CONT = 0
    While Not Rs.EOF
        'BUSCAMOS DATOS
        cad = "SELECT * from cuentas where codmacta='" & Rs.Fields(0) & "'"
    
        'Insertar datos en z347
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'Nuevo. Ya no llevamos NIF, llevaremos departamento
        RC = "" 'SERA EL NIF. Sera el DPTO
        I = 1
        If Not miRsAux.EOF Then
            'NIF -> codmacta
            RC = Rs.Fields(0)
            Dpto = Rs.Fields(1)
        Else
            'EOF
            I = 0
            MsgBox "No se encuentra la cuenta: " & Rs.Fields(0), vbExclamation
            'NOS SALIMOS
            Rs.Close
            Exit Sub
        End If
        
        'NO es EOF y tiene NIF
        If I > 0 Then
            'Aumentamos el contador
            CONT = CONT + 1
            
            
            'INSERTAMOS EN z347
            '-----------------------------------------
            Sql = "INSERT INTO Usuarios.z347 (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla,Provincia) "
            'Febrero 2010
            'SQL = SQL & "VALUES (" & vUsu.Codigo & ",0,'" & RC & "',0,'"
            Sql = Sql & "VALUES (" & vUsu.Codigo & "," & Dpto & ",'" & RC & "',0,'"
            
            
            'Razon social, dirdatos,codposta,despobla
            Sql = Sql & DevNombreSQL(DBLet(miRsAux!razosoci, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!dirdatos, "T")) & "','" & DBLet(miRsAux!codposta, "T") & "','" & DevNombreSQL(DBLet(miRsAux!desPobla, "T"))
            Sql = Sql & "','" & DevNombreSQL(DBLet(miRsAux!desProvi, "T"))
            Sql = Sql & "')"
        
            Conn.Execute Sql
        
        
        
            
            Sql = "INSERT INTO Usuarios.zcuentas (codusu, codmacta, nommacta,despobla,razosoci,dpto) VALUES (" & vUsu.Codigo & ",'" & RC & "','"
            Sql = Sql & DBLet(miRsAux!nifdatos, "T") & "','" 'En nommacta meto el NIF del cliente
            If IsNull(miRsAux!Entidad) Then
                'Puede que sean todos nulos
                cad = DBLet(miRsAux!Oficina) & "   " & DBLet(miRsAux!CC, "T") & "    " & DBLet(miRsAux!Cuentaba, "T")
                cad = Trim(cad)
            Else
                cad = DBLet(miRsAux!IBAN, "T") & " " & Format(miRsAux!Entidad, "0000") & " " & Format(DBLet(miRsAux!Oficina, "N"), "0000") & "  " & Format(DBLet(miRsAux!CC, "N"), "00") & " " & Format(DBLet(miRsAux!Cuentaba, "N"), "0000000000")
            End If
            cad = cad & "','"
            'El dpto si tiene
            cad = cad & DevNombreSQL(DevuelveDesdeBD("descripcion", "departamentos", "codmacta = '" & miRsAux!codmacta & "' AND dpto", CStr(Dpto)))
            cad = cad & "'," & Dpto
            Ejecuta Sql & cad & ")"   'Lo pongo en funcion para que no me de error
            
            
            'Updatear  FALTA### codusu = vusu.codusu
            Sql = "UPDATE USUARIOS.zentrefechas SET nomconam='" & RC & "' WHERE fechaadq = '" & Rs!fechaadq & "'"
            Sql = Sql & " AND codusu = " & vUsu.Codigo
            Conn.Execute Sql
            
            
            
        End If
        miRsAux.Close
            
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    
        
    If CONT = 0 Then
        MsgBox "Ningun dato devuelto para procesar por carta/mail", vbExclamation
        Exit Sub
    End If
    
    'Noviembre 2014
    'Comprobamos que todas las cuentas tienen email(si va por email)
    If Me.chkEmail.Value = 1 Then
            CadenaDesdeOtroForm = ""
            frmVarios.Opcion = 31
            frmVarios.Show vbModal
            
            If CadenaDesdeOtroForm = "" Then
                Screen.MousePointer = vbDefault
                Set Rs = Nothing
                Exit Sub
            End If
    End If
    'AHORA YA ESTA. Si es carta, imprimimios directamente
    If chkEmail.Value = 0 Then
        'POR CARTA
        cad = "FechaIMP= """ & Text3(8).Text & """|"
        cad = cad & "verCCC= " & Abs(Me.chkMostrarCta) & "|"
        CadenaDesdeOtroForm = NomArchivo
        
        With frmImprimir
            .EnvioEMail = False
            .OtrosParametros = cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .Opcion = 7
            '
            .Show vbModal
        End With


    Else
        'POR MAIL. IREMOS UNO A UNO
        ' fechaadq = codmacta
        Screen.MousePointer = vbHourglass
        
        cad = "DELETE FROM tmp347 WHERE codusu =" & vUsu.Codigo
        Conn.Execute cad
        
        cad = "SELECT fechaadq,maidatos,razosoci,nommacta FROM USUARIOS.zentrefechas,cuentas WHERE"
        cad = cad & " fechaadq=codmacta AND    CodUsu = " & vUsu.Codigo
        cad = cad & " GROUP BY fechaadq ORDER BY maidatos"
        Rs.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        cad = "FechaIMP= """ & Text3(8).Text & """|"
        cad = cad & "verCCC= " & Abs(Me.chkMostrarCta) & "|"
        Sql = "{ado.codusu}=" & vUsu.Codigo
        NumRegElim = 0
        CONT = 0
        frmPpal.Visible = False

        While Not Rs.EOF
            Me.Refresh
            espera 0.5
            RC = DBLet(Rs!maidatos, "T")
            If RC = "" Then
                
                If MsgBox("Sin mail para la cuenta: " & Rs!fechaadq & " - " & Rs!Nommacta & vbCrLf & "    ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then
                    CONT = 0
                    Rs.MoveLast
                End If
                
                Sql = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe) VALUES (" & vUsu.Codigo
                Sql = Sql & ",0," & Rs!fechaadq & ",NULL,0)"
                '
                'AL meter la cuenta con el importe a 0, entonces no la leera para enviarala
                'Pero despues si k podremos NO actualizar sus pagosya que no se han enviado nada
                Conn.Execute Sql
            Else
                Screen.MousePointer = vbHourglass
                With frmImprimir
                    CadenaDesdeOtroForm = NomArchivo
                    .OtrosParametros = cad
                    .NumeroParametros = 1
                    Sql = "{ado.codusu}=" & vUsu.Codigo & " AND {ado.nif}= """ & Rs.Fields(0) & """"
                    .FormulaSeleccion = Sql
                    .EnvioEMail = True
                    .QueEmpresaEs = EmpresaEscalona
                    .Opcion = 7
                    .Show vbModal
                    
                    If CadenaDesdeOtroForm = "OK" Then
                        Me.Refresh
                        espera 0.5
                        CONT = CONT + 1
                        'Se ha generado bien el documento
                        'Lo copiamos sobre app.path & \temp
                        Sql = Rs.Fields(0) & ".pdf"
                        
                        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Sql
                        
                        
                        'Insertamos en tmp347 la cuenta
                        Sql = "INSERT INTO tmp347(codusu, cliprov, cta, nif) VALUES (" & vUsu.Codigo & ",0,'" & Rs.Fields(0) & "','" & Sql & "')"
                        Conn.Execute Sql
                        
                    End If
                    
                End With
            End If
            Rs.MoveNext
        Wend
        Rs.Close

        If CONT > 0 Then
             
             espera 0.5
             
             Sql = "Reclamacion fecha: " & Text3(8).Text & "|"
             
             Sql = Sql & "Reclamación pago facturas efectuada el : " & Text3(8).Text & "|"
             
             'Escalona
             Sql = txtVarios(0).Text & "|Recuerde: En el archivo adjunto le enviamos información de su interés.|"
'--monica
'             frmEMail.QueEmpresa = EmpresaEscalona
'             frmEMail.Opcion = 3
'             frmEMail.MisDatos = SQL
'             frmEMail.Show vbModal
            
        End If
        
    End If
    
    Me.Hide
    frmPpal.Visible = True
    Me.Visible = True
    Me.Refresh
    
    Screen.MousePointer = vbHourglass
    
    'AHORA UPDATEAMOS LA FECHA RECLAMACION EN EL PAGO
    'SI ASI LO DESEA EL RECLAMANTE
    'Y SI SE HA REALIZADO; CUANTO MENOS; EL ENVIO
    '-----------------------------------------------------
    '-----------------------------------------------------
    If chkMarcarUtlRecla.Value = 1 Then
    
        'Si es por carta son todas, si es por mail, veremos si se ha llegado a enviar por mail, por lo menos
        'El mail sabemos k se ha enviado por que seran los k queden en tmp437
        'sin eliminar
        
        
        
        'Entonces veremos las reclamaciones k hemos efectuado bien, por email
        If Me.chkEmail.Value = 1 Then
            Sql = "SELECT * FROM tmp347 WHERE codusu=" & vUsu.Codigo & " AND Importe =0 "
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                'YA tengo la cuenta k no he podido enviar
                Sql = "DELETE from Usuarios.zentrefechas where codusu=" & vUsu.Codigo
                Sql = Sql & " AND nomconam = '" & Rs!Cta & "'"
                Conn.Execute Sql
                'Siguiente
                Rs.MoveNext
            Wend
            Rs.Close
        End If
        
            
        'AHORA, las que queden en entrefechas seran las k he enviado por mail, con lo cual
        ' el proceso es el mismo k el de cartas
        
        Sql = "SELECT * from Usuarios.zentrefechas where codusu = " & vUsu.Codigo
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = "UPDATE scobro set ultimareclamacion = '" & Format(Text3(8).Text, FormatoFecha) & "' WHERE numserie = '"
        While Not Rs.EOF
            'VAMOS A MARCAR EL PAGO CON LA FECHA UTLMARECLAMCION
            cad = Rs!codccost & "' AND codfaccl = " & Rs!nomccost & " AND fecfaccl  ='"
            cad = cad & Format(Rs!fecventa, FormatoFecha) & "' AND numorden = " & Rs!conconam
            cad = Sql & cad
            Conn.Execute cad
    
            'Siguiente
            Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    
    'FINALMENTE GRABAMOS LA TABLA HCO
    If chkInsertarReclamas.Value = 1 Then
        Sql = "SELECT MAX(codigo) from shcocob"
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CONT = 0
        If Not Rs.EOF Then CONT = DBLet(Rs.Fields(0), "N")
        Rs.Close
        CONT = CONT + 1
    
        'INSERT INTO shcocob (codigo, numserie, codfaccl, fecfaccl, numorden, impvenci, codmacta, nommacta, carta) VALUES (
        Sql = "SELECT zentrefechas.*,nommacta from Usuarios.zentrefechas,cuentas where codusu = " & vUsu.Codigo
        Sql = Sql & " AND zentrefechas.nomconam=cuentas.codmacta"
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = "INSERT INTO shcocob (fecreclama,carta,codigo, numserie, codfaccl, fecfaccl, numorden, impvenci, codmacta, nommacta)"
        Sql = Sql & " VALUES ('" & Format(Text3(8).Text, FormatoFecha) & "',"
        If Me.chkEmail.Value = 1 Then
            Sql = Sql & "1,"
        Else
            Sql = Sql & "0,"
        End If
        While Not Rs.EOF
            cad = CONT & ",'" & Rs!codccost & "'," & Rs!nomccost & ",'" & Format(Rs!fecventa, FormatoFecha) & "',"
            Importe = Rs!impventa + Rs!valoradq - DBLet(Rs!impperiodo, "N")
            cad = cad & Rs!conconam & "," & TransformaComasPuntos(CStr(Importe)) & ",'"
            cad = cad & Rs!nomconam & "','" & DevNombreSQL(Rs!Nommacta) & "')"
            cad = Sql & cad
            Conn.Execute cad
            'siguiente
            CONT = CONT + 1
            Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub MontaSQLReclamacion()
    
    'Siempre hay que añadir el AND
    
    
    Sql = ""
    
    
    'Fecha factura
    RC = CampoABD(txtSerie(2), "T", "scobro.numserie", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(txtSerie(3), "T", "scobro.numserie", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    
    'Fecha factura
    RC = CampoABD(Text3(6), "F", "fecfaccl", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    
    RC = CampoABD(Text3(7), "F", "fecfaccl", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    'Fecha vto
    RC = CampoABD(Text3(9), "F", "fecvenci", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(Text3(10), "F", "fecvenci", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    'cuenta
    RC = CampoABD(txtCta(4), "T", "scobro.codmacta", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(txtCta(5), "T", "scobro.codmacta", False)
    If RC <> "" Then Sql = Sql & " AND " & RC

    
    
    'Agente
    RC = CampoABD(txtAgente(3), "N", "scobro.agente", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(txtAgente(2), "N", "scobro.agente", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    
    'Forma de pago
    RC = CampoABD(txtFPago(3), "N", "scobro.codforpa", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(txtFPago(2), "N", "scobro.codforpa", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    'Solo devueltos
    If chkReclamaDevueltos.Value = 1 Then Sql = Sql & " AND devuelto = 1"
      
    
    'Marzo2015
    If chkExcluirConEmail.Value = 1 Then Sql = Sql & " AND coalesce(maidatos,'')=''"
    
    
    'LA de la fecha
    Sql = Sql & " AND ((ultimareclamacion  is null) OR (ultimareclamacion <= '" & Format(Fecha, FormatoFecha) & "'))"
    
    'QUE FALTE POR PAGAR
    Sql = Sql & " AND (impvenci>0)"
    
    
    RC = PonerTipoPagoCobro_(True, True)
    If RC <> "" Then Sql = Sql & " AND tipforpa IN " & RC
    
    
    
    'Select
    cad = "Select scobro.*, cuentas.codmacta FROM scobro,cuentas,sforpa "
    cad = cad & " WHERE  sforpa.codforpa=scobro.codforpa AND scobro.codmacta = cuentas.codmacta"
    cad = cad & " AND sforpa.codforpa=scobro.codforpa "
    Sql = cad & Sql
    
    
    
    
    
End Sub






Private Sub cmdReclamas_Click()
    
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If ListadoReclamas Then
        With frmImprimir
            cad = "Cadena= """ & cad & """|"
            .OtrosParametros = cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 86
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdTransfer_Click()
    
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If ListadoTransferencias Then
        With frmImprimir
            cad = "Mostrar= 1|tipot= """
            I = 28
            Sql = "Listado transferencias"
            If Opcion = 11 Then
                cad = cad & "(Pagos)"
            ElseIf Opcion = 13 Then
                cad = cad & "(Abonos)"
                If Me.chkCartaAbonos.Value Then
                    CadenaDesdeOtroForm = DevuelveNombreInformeSCRYST(12, "Carta abono0")
                    I = 95
                End If
            Else
                If Opcion = 44 Then
                    Sql = "Caixa confirming"
                Else
                    Sql = "Pagos domiciliados"
                End If
            End If
            
            cad = cad & """|ErTitulo= """ & Sql & """|"
            
            .OtrosParametros = cad
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            
            .Opcion = I
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVtoDestino_Click(Index As Integer)
    
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
    
        cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
        If cad = "" Then
            MsgBox "No esta configurada la aplicación. Falta informe(10)", vbCritical
            Exit Sub
        End If
        CadenaDesdeOtroForm = cad
    
        LanzaBuscaGrid 1, 4


    End If
End Sub

Private Sub Command1_Click()
    If txtCta(13).Text = "" Then
        MsgBox "Ponga la cuenta", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = txtCta(13).Text
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 1
            Text3(1).SetFocus
        Case 3
            
            'Reclamaciones. Si no tiene configurado el envio web
            'no habilitaremos el check
            cad = DevuelveDesdeBD("smtpHost", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
            If cad = "" Then
                Me.chkEmail.Value = 0
                chkEmail.Enabled = False
            End If
            'Text3(6).SetFocus
            txtSerie(2).SetFocus
        Case 10
            Me.cmdFormaPago.SetFocus
        Case 12
            txtCtaBanc(0).SetFocus
        Case 20
            PonFoco txtCta(13)
            
        Case 22
            'Contabi efectos
            If CONT > 0 Then
                For I = 1 To Me.cboCompensaVto.ListCount
                    If Me.cboCompensaVto.ItemData(I) = CONT Then
                        CONT = I
                        Exit For
                    End If
                Next
            End If
            Me.cboCompensaVto.ListIndex = CONT
            PonFoco Text3(23)
        Case 23
            CadenaDesdeOtroForm = ""  'Para que  no devuelva nada
        Case 30
            PonFoco Text3(28)
            
        Case 31
            'gastos fijos
            Text3(30).Text = "01/01/" & Year(Now)
        Case 35
            PonFoco txtImporte(2)
            
        Case 36
            If CadenaDesdeOtroForm <> "" Then
                txtCta(17).Text = CadenaDesdeOtroForm
                txtCta_LostFocus 17
            Else
                PonFoco txtCta(17)
            End If
            CadenaDesdeOtroForm = ""
            
        Case 39
            PonFoco Text3(34)
            
        Case 42
            
'            Me.Refresh
'            cmdNoram57Fich_Click
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub



    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmPpal.Icon
    CargaImagenesAyudas Me.imgCtaBanc, 1, "Cuenta contable bancaria"
    CargaImagenesAyudas Image2, 2
    CargaImagenesAyudas Me.imgFP, 1, "Forma de pago"
    CargaImagenesAyudas Me.Image3, 1, "Cuenta contable"
    CargaImagenesAyudas Me.Imagente, 1, "Seleccionar agente"
    CargaImagenesAyudas imgCCoste, 1, "Centro de coste"
    CargaImagenesAyudas Me.ImageAyudaImpcta, 3
    For Each Img In Me.imgConcepto
        Img.ToolTipText = "Concepto"
    Next
    For Each Img In Me.imgDiario
        Img.ToolTipText = "Diario"
    Next
    
    
    
    For Each Img In Me.imgDpto
        Img.ToolTipText = "Departamento"
    Next
    
    
    'Limpiamos el tag
    txtCta(6).Tag = ""
    PrimeraVez = True
    FrCobrosPendientesCli.Visible = False
    frpagosPendientes.Visible = False
    FramereclaMail.Visible = False
    FrameAgentes.Visible = False
    FrameDpto.Visible = False
    FrameListRem.Visible = False
    FrameListadoCaja.Visible = False
    FrameDevEfec.Visible = False
    Me.FrameFormaPago.Visible = False
    FrameTransferencias.Visible = False
    Me.FramePrevision.Visible = False
    FrameAseg_Bas.Visible = False
    FrameCobroGenerico.Visible = False
    FrameCompensaciones.Visible = False
    FrameRecepcionDocumentos.Visible = False
    FrameListaRecep.Visible = False
    frameListadoPagosBanco.Visible = False
    FrameDividVto.Visible = False
    FrameReclama.Visible = False
    FrameGastosFijos.Visible = False
    FrameGastosTranasferencia.Visible = False
    FrameCompensaAbonosCliente.Visible = False
    FrameRecaudaEjec.Visible = False
    FrameOperAsegComunica.Visible = False
    FrameNorma57Importar.Visible = False
    CommitConexion
    
    Select Case Opcion
    Case 1
        'Leeo la opcion del fichero x defecto
        Me.optCuenta(0).Value = CheckValueLeer("Listcta") = 1
        If Me.optCuenta(0).Value = False Then Me.optCuenta(1).Value = True
    
        chkApaisado(0).Value = Abs(CheckValueLeer("Infapa") = 1)
    
        Me.Frame1.BorderStyle = 0 'sin borde
        FrCobrosPendientesCli.Visible = True
        W = Me.FrCobrosPendientesCli.Width
        H = Me.FrCobrosPendientesCli.Height + 120
        Text3(0).Text = Format(Now, "dd/mm/yyyy")
        'Fecha = CDate(DiasMes(Month(Now), Year(Now)) & "/" & Month(Now) & "/" & Year(Now))
        'Text3(2).Text = Format(Fecha, "dd/mm/yyyy")
        Me.cmbCuentas(0).Tag = ""
        GeneraComboCuentas
        Me.cmbCuentas(0).ListIndex = 0
        
        Me.cboCobro(0).ListIndex = 2
        Me.cboCobro(1).ListIndex = 0
    Case 2
        frpagosPendientes.Visible = True
        W = Me.frpagosPendientes.Width
        H = Me.frpagosPendientes.Height
        Text3(5).Text = Format(Now, "dd/mm/yyyy")
        'Fecha = CDate(DiasMes(Month(Now), Year(Now)) & "/" & Month(Now) & "/" & Year(Now))
        'Text3(4).Text = Format(Fecha, "dd/mm/yyyy")
        Me.cmbCuentas(1).Tag = ""
        GeneraComboCuentas
        Me.cmbCuentas(1).ListIndex = 0
    Case 3
        Caption = "Reclamaciones"
        FramereclaMail.Visible = True
        W = Me.FramereclaMail.Width
        H = Me.FramereclaMail.Height
        Text3(8).Text = Format(Now, "dd/mm/yyyy")

        'ESPECIAL
        'Si no existe la carpeta tmp en app.path la creo
        If Dir(App.Path & "\temp", vbDirectory) = "" Then MkDir App.Path & "\temp"
        CargaTextosTipoPagos True
    Case 4
        
        Caption = "Agentes"
        FrameAgentes.Visible = True
        W = Me.FrameAgentes.Width
        H = Me.FrameAgentes.Height
        
    Case 5
         
        Caption = "Departamentos"
        FrameDpto.Visible = True
        W = Me.FrameDpto.Width
        H = Me.FrameDpto.Height
        
        
    Case 6, 7
         
        Caption = "Remesas"
        FrameListRem.Visible = True
        W = Me.FrameListRem.Width
        H = Me.FrameListRem.Height
        FrameOrdenRemesa.Visible = False
        
    Case 8
        FrameListadoCaja.Visible = True
        Caption = "Listado"
        W = Me.FrameListadoCaja.Width
        H = Me.FrameListadoCaja.Height
        
        
    Case 9
        
        FrameDevEfec.Visible = True
        Caption = "Listado"
        W = Me.FrameDevEfec.Width
        H = Me.FrameDevEfec.Height + 120
        
        
    Case 10
        
        FrameFormaPago.Visible = True
        Caption = "Listado"
        W = Me.FrameFormaPago.Width
        H = Me.FrameFormaPago.Height
    Case 11, 13, 43, 44
        
        FrameTransferencias.Visible = True
        
        If Opcion < 43 Then
            Label2(9).Caption = "Listado transferencias"
            Sql = "Listado trans."
            If Opcion = 11 Then
                'Puede ser transferencias o confirmings
                Caption = "PROVEEDORES"
            Else
                Caption = "ABONOS"
            End If
        
        Else
            Sql = "Listado "
            If Opcion = 43 Then
                'Puede ser transferencias o confirmings
                Caption = "Pagos domiciliados"
            Else
                Caption = "Caixa confirming"
            End If
            Label2(9).Caption = Caption
        End If
        Caption = Sql & " " & Caption
        W = Me.FrameTransferencias.Width
        H = Me.FrameTransferencias.Height + 60
        chkCartaAbonos.Visible = Opcion = 13
        
    Case 12
        
        FramePrevision.Visible = True
        Caption = "Listado"
        W = Me.FramePrevision.Width
        H = Me.FramePrevision.Height
        Text3(18).Text = Format(DateAdd("m", 2, Now), "dd/mm/yyyy")
        
        
    Case 15, 16, 17, 18, 33
        
        'Operaciones aseguradas
        '       Datos basicos
        '       Listado facturacion
        '       Impagados
        optAsegBasic(2).Visible = True 'Ordenar por poliza
        FrOrdenAseg1.Visible = True
        FrameASeg2.Visible = False
        FrameForpa.Visible = False
        FrameAsegAvisos.Visible = False
        Select Case Opcion
        Case 15
            '       Datos basicos
            Sql = "Fecha solicitud"
            cad = "Datos básicos operaciones aseguradas"
            
        Case 16
            '       Listado facturacion
            Sql = "Fecha"
            cad = "List. facturacion oper. aseguradas"
            FrOrdenAseg1.Visible = False
            FrameASeg2.Visible = True
            FrameForpa.Visible = True
        Case 17
            '       Listado impagados asegurados
            Sql = "Fecha aviso"
            cad = "Impagados en operaciones aseguradas"
            
        Case 18
            optAsegBasic(2).Visible = False
            Sql = "Fecha vto"
            cad = "Listado efectos operaciones aseguradas"
            
        Case 33
            FrameAsegAvisos.Visible = True
           ' FrOrdenAseg1.Visible = False
            Sql = "Fecha aviso falta pago"
            cad = "Listados avisos aseguradoras"
            optAsegAvisos(0).Value = True
        End Select
        
        
        Label4(39).Caption = Sql
        Label2(11).Caption = cad
        FrameAseg_Bas.Visible = True
        Caption = "Listado"
        W = Me.FrameAseg_Bas.Width
        H = Me.FrameAseg_Bas.Height
        
        
        
    Case 20
        H = FrameCobroGenerico.Height + 120
        W = FrameCobroGenerico.Width
        FrameCobroGenerico.Visible = True
        Caption = "Cuenta"
    Case 22
        
        
        For H = 0 To 1
            
            txtConcpto(H).Text = RecuperaValor(CadenaDesdeOtroForm, (H * 2) + 1)
            txtDescConcepto(H).Text = RecuperaValor(CadenaDesdeOtroForm, (H * 2) + 2)
        Next H
        Me.cboCompensaVto.Clear
        InsertaItemComboCompensaVto "No compensa sobre ningún vencimiento", 0
        
        'Veremos si puede sobre un Vto o no
        H = RecuperaValor(CadenaDesdeOtroForm, 5)
        CONT = 0
        If H = 1 Then CONT = RecuperaValor(CadenaDesdeOtroForm, 6)
        FrameCambioFPCompensa.Visible = CONT > 0
        'chkCompensaVto.Value = 0
        'chkCompensaVto.Enabled = h = 1
        'chkCompensaVto.Caption = RecuperaValor(CadenaDesdeOtroForm, 6)
        CadenaDesdeOtroForm = ""
        H = FrameCompensaciones.Height + 120
        W = FrameCompensaciones.Width
        FrameCompensaciones.Visible = True
        Caption = "Compensacion efectos"
        Text3(23).Text = Format(Now, "dd/mm/yyyy")
        
        
    Case 23, 34
        '23.-  Contabilizar
        '34. Eliminar ya contabilizada
        
        
        
        
        'Tendremos el tipo de pago , talon o pagare
        Dim FP As Ctipoformapago
        Set FP = New Ctipoformapago
        
        If Opcion = 23 Then
            Label2(13).Caption = "Contabilizar recepción documentos"
            Caption = "Contabilizar"
        Else
            Label2(13).Caption = "Eliminar de recepción documentos"
            Caption = "Eliminar"
        End If
        
        'Cuenta beneficios gastos paras las diferencias si existieran
        'Si el total del talon es el total de las lineas entonces no mostrara los
        'datos del total. 0: igual   1  Mayor     2 Menor
        Sql = RecuperaValor(CadenaDesdeOtroForm, 2)
        I = CInt(Sql)
'        If CInt(SQL) > 0 Then
'            I = 1
'        Else
'            I = -1
'        End If
        
        Label4(55).Visible = I <> 0
        Image3(14).Visible = I <> 0
        txtCta(14).Visible = I <> 0
        DtxtCta(14).Visible = I <> 0
        Label6(28).Visible = I <> 0
        
        
        
        
        
        If I > 0 Then
            Sql = "Beneficios"
        Else
            Sql = "Pérdidas"
        End If
        
        If Opcion = 34 Then Sql = Sql & "(Deshacer apunte)"
        Label4(55).Caption = Sql

        


        '   No lleva ANALITICA
        If I <> 0 Then
            If Not vParam.autocoste Then I = 0
        End If
     
        Me.imgCCoste(0).Visible = I <> 0
        Me.txtCCost(0).Visible = I <> 0
        Label6(29).Visible = I <> 0
        Me.txtDescCCoste(0).Visible = I <> 0
     
        
        
        
        
        
        
        
        
        Sql = RecuperaValor(CadenaDesdeOtroForm, 1)
        I = CInt(Sql)
        If FP.Leer(I) = 0 Then
            If Opcion = 23 Then
                'Normal
                txtDiario(1).Text = FP.diaricli
                txtConcpto(2).Text = FP.condecli
                txtConcpto(3).Text = FP.conhacli
             Else
                'Eliminar. Iran cambiados
                txtDiario(1).Text = FP.diaricli
                txtConcpto(2).Text = FP.conhacli
                txtConcpto(3).Text = FP.condecli
                
                
             End If
                
            'Para que pinte la descripcion
            txtDiario_LostFocus 1
            txtConcpto_LostFocus 2
            txtConcpto_LostFocus 3
        End If
        
        
        
        
        H = 0
        If I = vbTalon Then
            Sql = "taloncta"
        Else
            Sql = "pagarecta"
        End If
        
        Sql = DevuelveDesdeBD(Sql, "paramtesor", "codigo", "1")
        If Len(Sql) = vEmpresa.DigitosUltimoNivel Then
            chkAgruparCtaPuente(0).Visible = True
            H = 1 '
        
            'Si esta configurado en parametrps, si la ultima vez lo marco seguira marcado
            If H = 1 Then H = CheckValueLeer("Agrup0")
            If H <> 1 Then H = 0
            chkAgruparCtaPuente(0).Value = H
            
        Else
            chkAgruparCtaPuente(0).Visible = False
        End If
        
        Set FP = Nothing
        
        If Label4(55).Visible Then '5055
            FrameRecepcionDocumentos.Height = 4815
            I = 4320
        Else
            FrameRecepcionDocumentos.Height = 3135
            I = 2640
        End If
        cmdRecepDocu.Top = I
        cmdCancelar(23).Top = I
        H = FrameRecepcionDocumentos.Height + 120
        W = FrameRecepcionDocumentos.Width
        FrameRecepcionDocumentos.Visible = True
        
        
            
        
        
        
    Case 24
        
        H = FrameListaRecep.Height + 120
        W = FrameListaRecep.Width
        FrameListaRecep.Visible = True
        
        
    Case 25
        
                
        H = frameListadoPagosBanco.Height + 120
        W = frameListadoPagosBanco.Width
        frameListadoPagosBanco.Visible = True
        
        
    Case 26
        'Si el total del talon es el total de las lineas entonces no mostrara los
        'datos del total. 0: igual   1  Mayor     2 Menor
        Sql = RecuperaValor(CadenaDesdeOtroForm, 1)
        If CCur(Sql) > 0 Then
            I = 1
        Else
            I = -1
        End If
        
        'Label4(55).Visible = True
        'Image3(14).Visible = True
        'txtCta(14).Visible = True
        'DtxtCta(14).Visible = I <> 0
        'Label6(28).Visible = I <> 0
        
         'If I > 0 Then
         '    SQL = "Beneficios"
         'Else
         '    SQL = "Pérdidas"
         'End If
         'Label4(55).Caption = SQL


        '   No lleva ANALITICA
        I = 1
        If Not vParam.autocoste Then I = 0
     
        Me.imgCCoste(0).Visible = I <> 0
        Me.txtCCost(0).Visible = I <> 0
        Label6(29).Visible = I <> 0
        Me.txtDescCCoste(0).Visible = I <> 0
        If I <> 0 Then CargaImagenesAyudas imgCCoste(0), 1 '--++monica Carga1ImagenAyuda
        
        
'        h = FrameCancelRemTalPag.Height + 120
'        W = FrameCancelRemTalPag.Width
'        FrameCancelRemTalPag.Visible = True
        
        
    Case 27
                'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
        H = FrameDividVto.Height + 120
        W = FrameDividVto.Width
        FrameDividVto.Visible = True
        
    Case 30
        H = FrameReclama.Height + 120
        W = FrameReclama.Width
        FrameReclama.Visible = True
        
    Case 31
        H = FrameGastosFijos.Height + 120
        W = FrameGastosFijos.Width
        FrameGastosFijos.Top = 0
        FrameGastosFijos.Left = 90
        FrameGastosFijos.Visible = True
        
    Case 35
        Me.txtVarios(2).Text = CadenaDesdeOtroForm
        CadenaDesdeOtroForm = ""
        H = FrameGastosTranasferencia.Height + 120
        W = FrameGastosTranasferencia.Width
        FrameGastosTranasferencia.Visible = True
        
        
    Case 36
        
        
        H = FrameCompensaAbonosCliente.Height + 120
        W = FrameCompensaAbonosCliente.Width
        FrameCompensaAbonosCliente.Visible = True
        
        
        'cmdVtoDestino(1).Visible = (vUsu.Codigo Mod 100) = 0
        'Label1(1).Visible = (vUsu.Codigo Mod 100) = 0
        cmdVtoDestino(1).Visible = vUsu.Nivel = 0
        Label1(1).Visible = vUsu.Nivel = 0
        
        
    Case 38
        
        H = FrameRecaudaEjec.Height + 120
        W = FrameRecaudaEjec.Width
        FrameRecaudaEjec.Visible = True
        Fecha = DateAdd("yyyy", -4, Now)
        Text3(32).Text = Format(Fecha, "dd/mm/yyyy")
        
        
    Case 39, 40
        H = FrameOperAsegComunica.Height + 120
        W = FrameOperAsegComunica.Width
        FrameOperAsegComunica.Visible = True
        
        cargaEmpresasTesor ListView3
        Fecha = Now
        If Day(Now) < 15 Then Fecha = DateAdd("m", -1, Now)
        
        Text3(34).Text = "01/" & Format(Fecha, "mm/yyyy")
        
            
        Text3(35).Text = Format(Now, "dd/mm/yyyy")
        
        FrameSelEmpre1.BorderStyle = 0
        FrameFraPendOpAseg.BorderStyle = 0
        
        FrameSelEmpre1.Visible = Opcion = 39
        FrameFraPendOpAseg.Visible = Opcion = 40
        
        
        If Opcion = 39 Then
            Label2(22).Caption = "Comunicación datos al seguro"
        Else
            Label2(22).Caption = "Fras. pendientes op. aseguradas"
        End If
        
        
    Case 42
        H = FrameNorma57Importar.Height + 120
        W = FrameNorma57Importar.Width
        FrameNorma57Importar.Visible = True
    
    
        
        
    End Select
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    I = Opcion
    If Opcion = 13 Or I = 43 Or I = 44 Then I = 11
    
    'Aseguradas
    If Opcion >= 15 And Opcion <= 18 Then I = 15  'aseguradoas
    If Opcion = 33 Then I = 15 'aseguradoas
    If Opcion = 34 Then I = 23 'Eliminar recepcion documento
    If Opcion = 40 Then I = 39
    Me.cmdCancelar(I).Cancel = True
    
    PonerFrameProgreso

End Sub


Private Sub PonerFrameProgreso()
Dim I As Integer

    'Ponemos el frame al pricnipio de todo
    FrameProgreso.Visible = False
    FrameProgreso.ZOrder 0
    
    'lo ubicamos
    'Posicion horizintal WIDTH
    I = Me.Width - FrameProgreso.Width
    If I > 100 Then
        I = I \ 2
    Else
        I = 0
    End If
    FrameProgreso.Left = I
    'Posicion  VERTICAL HEIGHT
    I = Me.Height - FrameProgreso.Height
    If I > 100 Then
        I = I \ 2
    Else
        I = 0
    End If
    FrameProgreso.Top = I
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If Opcion = 1 Then
        CheckValueGuardar "Listcta", CByte(Abs(Me.optCuenta(0).Value))
        CheckValueGuardar "Infapa", chkApaisado(0)
    End If
    If Opcion = 23 Then CheckValueGuardar "Agrup0", Me.chkAgruparCtaPuente(0)
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtAgente(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescAgente(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    DevfrmCCtas = CadenaDevuelta
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub







Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    'Si no habia cuenta
    
        txtCta(0).Text = RecuperaValor(CadenaSeleccion, 1)
        DtxtCta(0).Text = RecuperaValor(CadenaSeleccion, 2)
        txtCta(1).Text = RecuperaValor(CadenaSeleccion, 1)
        DtxtCta(1).Text = RecuperaValor(CadenaSeleccion, 2)
    
    'El dpto
    txtDpto(RC).Text = RecuperaValor(CadenaSeleccion, 3)
    txtDescDpto(RC).Text = RecuperaValor(CadenaSeleccion, 4)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtFPago(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescFPago(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    txtSerie(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing
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










Private Sub ImageAyudaImpcta_Click(Index As Integer)
Dim C As String
    Select Case Index
    Case 0
            C = "Compensaciones" & vbCrLf & String(60, "-") & vbCrLf
            C = C & "Cuando compense sobre un vencimiento al marcar la opción " & vbCrLf
            C = C & Space(10) & Me.chkCompensa.Caption & vbCrLf
            C = C & "se modificará el importe vencimiento poniendo el total a compensar  y en importe cobrado un cero"
    Case 1
            C = "Asegurados" & vbCrLf & String(60, "-") & vbCrLf
            C = C & "Fecha 'hasta' es campo obligado para considerar la fecha de baja de los asegurados." & vbCrLf
            C = C & "En los listados saldrán aquellos que si tienen fecha de baja , es superior al hasta solicitado "
            
            'ALZIRA
            C = C & vbCrLf & vbCrLf & "Comunicación datos seguro" & vbCrLf & "Salen TODAS las facturas entre el periodo seleccionado para los"
            C = C & " clientes asegurados"
            
    End Select
    MsgBox C, vbInformation

End Sub

Private Sub Imagente_Click(Index As Integer)
    Set frmA = New frmAgentes
    RC = Index
    frmA.DatosADevolverBusqueda = "0|1|"
    frmA.Show vbModal
    Set frmA = Nothing
End Sub

Private Sub ImageSe_Click(Index As Integer)
    RC = Index
'    Set frmS = New frmSerie
'    frmS.DatosADevolverBusqueda = "0"
'    frmS.Show vbModal
    Set frmS = New frmBasico
    AyudaContadores frmS, txtSerie(RC).Text, "tiporegi REGEXP '^[0-9]+$' = 0"
    Set frmS = Nothing


    Set frmS = Nothing
End Sub

Private Sub imgCarta_Click()
    Screen.MousePointer = vbHourglass
'--monica
'    Set frmB = New frmBuscaGrid
'    DevfrmCCtas = ""
'    frmB.vSQL = ""
'
'    '###A mano
'    frmB.vDevuelve = "0|1|"   'Siempre el 0
'
'    frmB.vSelElem = 1
'
'    cad = "Codigo|codcarta|N|15·"
'    cad = cad & "Descripcion|descarta|T|65·"
'    frmB.vCampos = cad
'    frmB.vTabla = "scartas"
'    frmB.vTitulo = "Cartas reclamación"
'    frmB.Show vbModal
'    Set frmB = Nothing
'    If DevfrmCCtas <> "" Then
'        Me.txtDescCarta.Text = RecuperaValor(DevfrmCCtas, 2)
'        txtCarta.Text = RecuperaValor(DevfrmCCtas, 1)
'    End If
End Sub

Private Sub ImgCCoste_Click(Index As Integer)
    LanzaBuscaGrid Index, 2
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For I = 1 To Me.ListView3.ListItems.Count
        Me.ListView3.ListItems(I).Checked = (Index = 1)
    Next
        
End Sub

Private Sub imgConcepto_Click(Index As Integer)
    LanzaBuscaGrid Index, 1
End Sub

Private Sub imgCtaBanc_Click(Index As Integer)
    Sql = ""
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
    If Sql <> "" Then
        txtCtaBanc(Index).Text = RecuperaValor(Sql, 1)
        Me.txtDescBanc(Index).Text = RecuperaValor(Sql, 2)
    End If
End Sub

Private Sub imgDiario_Click(Index As Integer)
    LanzaBuscaGrid Index, 0
End Sub

Private Sub imgDpto_Click(Index As Integer)
    Sql = "NO"
    If txtCta(1).Text <> "" And txtCta(0).Text <> "" Then
        
        If txtCta(1).Text <> txtCta(0).Text Then
            MsgBox "Debe seleccionar un mismo cliente", vbExclamation
            txtDpto(Index).Text = ""
            Sql = ""
        End If
    End If
    If Sql = "" Then Exit Sub
'--monica
'    Set frmD = New frmDepartamentos
'    RC = Index
'    frmD.DatosADevolverBusqueda = "1|2|"
'    frmD.vCuenta = txtCta(0).Text
'    frmD.Show vbModal
    
    RC = Index
    
    Set frmD = New frmBasico
    AyudaDepartamentos frmD, txtDpto(RC).Text, "codmacta = " & DBSet(txtCta(1).Text, "T")
    Set frmD = Nothing
    PonFoco txtDpto(RC)
    
    
    
    
    Set frmD = Nothing
End Sub


Private Sub imgFP_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    'Set frmCta = New frmColCtas
    Set frmP = New frmFormaPago
    RC = Index
    frmP.DatosADevolverBusqueda = "0|1"
    frmP.Show vbModal
    Set frmP = Nothing
End Sub






Private Sub imgGastoFijo_Click(Index As Integer)
     LanzaBuscaGrid Index, 3
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

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

Private Sub optAsegAvisos_Click(Index As Integer)
    If Index = 0 Then
        Label4(39).Caption = "Fecha aviso falta pago"
    ElseIf Index = 1 Then
        Label4(39).Caption = "Fecha aviso prorroga"
    Else
        Label4(39).Caption = "Fecha aviso siniestro"
    End If
    
End Sub

Private Sub optAsegAvisos_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub optAsegBasic_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub optImpago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optLCobros_Click(Index As Integer)
    Me.Check1.Enabled = Me.optLCobros(1).Value
    Me.Check2.Enabled = Not Me.Check1.Enabled
End Sub

Private Sub optLCobros_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optPrevision_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optProv_Click(Index As Integer)
    Me.chkProv.Enabled = Me.optProv(1).Value
    Me.chkProv2.Enabled = Not Me.chkProv.Enabled
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















Private Sub txtAgente_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtAgente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAgente_LostFocus(Index As Integer)

    Sql = ""
    txtAgente(Index).Text = Trim(txtAgente(Index).Text)
    If txtAgente(Index).Text <> "" Then
        
        If Not IsNumeric(txtAgente(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtAgente(Index).Text = ""
            SubSetFocus txtAgente(Index)
        Else
            txtAgente(Index).Text = Val(txtAgente(Index).Text)
            Sql = DevuelveDesdeBD("nombre", "agentes", "codigo", txtAgente(Index).Text, "N")
            If Sql = "" Then Sql = "AGENTE NO ENCONTRADO"
        End If
    End If
    Me.txtDescAgente(Index).Text = Sql
        
End Sub





Private Sub txtCarta_GotFocus()
    PonFoco txtCarta
End Sub

Private Sub txtcarta_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCarta_LostFocus()
    Sql = ""
    txtCarta.Text = Trim(txtCarta.Text)
    If txtCarta.Text <> "" Then
        
        If Not IsNumeric(txtCarta.Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtCarta.Text = ""
            SubSetFocus txtCarta
        Else
            txtCarta.Text = Val(txtCarta.Text)
            Sql = DevuelveDesdeBD("descarta", "scartas", "codcarta", txtCarta.Text, "N")
            If Sql = "" Then txtCarta.Text = ""
        End If
    End If
    Me.txtDescCarta.Text = Sql
End Sub





Private Sub txtCCost_GotFocus(Index As Integer)
    PonFoco txtConcpto(Index)
End Sub

Private Sub txtCCost_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCCost_LostFocus(Index As Integer)
    Sql = ""
    txtCCost(Index).Text = Trim(txtCCost(Index).Text)
    If txtCCost(Index).Text <> "" Then
        

            txtConcpto(Index).Text = Val(txtConcpto(Index).Text)
            Sql = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtCCost(Index).Text, "T")
            If Sql = "" Then
                MsgBox "No existe el centro de coste: " & Me.txtCCost(Index).Text, vbExclamation
                Me.txtCCost(Index).Text = ""
            End If
        If txtCCost(Index).Text = "" Then SubSetFocus txtCCost(Index)
    End If
    Me.txtDescCCoste(Index).Text = Sql
End Sub

Private Sub txtConcpto_GotFocus(Index As Integer)
     PonFoco txtConcpto(Index)
End Sub

Private Sub txtConcpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtConcpto_LostFocus(Index As Integer)
    Sql = ""
    txtConcpto(Index).Text = Trim(txtConcpto(Index).Text)
    If txtConcpto(Index).Text <> "" Then
        
        If Not IsNumeric(txtConcpto(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtConcpto(Index).Text = ""
        Else
            txtConcpto(Index).Text = Val(txtConcpto(Index).Text)
            Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcpto(Index).Text, "N")
            If Sql = "" Then
                MsgBox "No existe el concepto: " & Me.txtConcpto(Index).Text, vbExclamation
                Me.txtConcpto(Index).Text = ""
            End If
        End If
        If txtConcpto(Index).Text = "" Then SubSetFocus txtConcpto(Index)
    End If
    Me.txtDescConcepto(Index).Text = Sql
    
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
    If Index = 6 Then
        'NO se ha cambiado nada de la cuenta
        If txtCta(6).Text = txtCta(6).Tag Then
        
            Exit Sub
        Else
            txtDpto(0).Text = ""
            txtDpto(1).Text = ""
            txtDescDpto(0).Text = ""
            txtDescDpto(0).Text = ""
        End If
    End If
     
     
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
       ' txtCta(6).Tag = txtCta(6).Text
        Exit Sub
    End If
    
    If Index = 6 Then
        If txtCta(0).Text <> "" Or txtCta(1).Text <> "" Then
            MsgBox "Si selecciona desde / hasta cliente no podra seleccionar departamento", vbExclamation
            txtCta(6).Text = ""
            txtCta(6).Tag = txtCta(6).Text
            Exit Sub
        End If
        
    Else
        If Index = 0 Or Index = 1 Then
            If txtCta(6).Text <> "" Then
                MsgBox "Si seleciona departamento no puede seleccionar desde / hasta  cliente", vbExclamation
                txtCta(Index).Text = ""
                txtCta(6).Tag = txtCta(6).Text
                Exit Sub
            End If
        End If
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
    Case 0 To 7, 11, 12, 15, 16, 18, 19
        'NO hace falta que sean de ultimo nivel
        Cta = (txtCta(Index).Text)
                                '********
        B = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
        If B = 0 Then
            MsgBox "NO existe la cuenta: " & txtCta(Index).Text, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
        Else
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = Sql
            If B = 1 Then
                DtxtCta(Index).Tag = ""
            Else
                DtxtCta(Index).Tag = Sql
            End If
            
            
            'Index=1. Cliente en listado de cobros. Si pongo el desde pongo el hasta lo mismo
            If Index = 1 Then
                
                If Len(Cta) = vEmpresa.DigitosUltimoNivel Then
                    txtCta(0).Text = Cta
                    DtxtCta(0).Text = DtxtCta(1).Text
                End If
            End If
            
        End If
    Case Else
        'DE ULTIMO NIVEL
        Cta = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(Cta, Sql) Then
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = Sql
            
            
        Else
            MsgBox Sql, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
            txtCta(Index).SetFocus
        End If
        If Index = 17 Then PonerVtosCompensacionCliente
        
    End Select
    txtCta(6).Tag = txtCta(6).Text
End Sub







Private Sub txtDiario_GotFocus(Index As Integer)
    PonFoco txtDiario(Index)
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
    
    Sql = ""
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    If txtDiario(Index).Text <> "" Then
        
        If Not IsNumeric(txtDiario(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtDiario(Index).Text = ""
            SubSetFocus txtDiario(Index)
        Else
            txtDiario(Index).Text = Val(txtDiario(Index).Text)
            Sql = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text, "N")
            
            If Sql = "" Then
                MsgBox "No existe el diario: " & Me.txtDiario(Index).Text, vbExclamation
                Me.txtDiario(Index).Text = ""
                PonFoco txtDiario(Index)
            End If
        End If
    End If
    Me.txtDescDiario(Index).Text = Sql
     
End Sub





Private Sub txtGastoFijo_GotFocus(Index As Integer)
    PonFoco txtGastoFijo(Index)
End Sub

Private Sub txtGastoFijo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtGastoFijo_LostFocus(Index As Integer)
    
    Sql = ""
    txtGastoFijo(Index).Text = Trim(txtGastoFijo(Index).Text)
    If txtGastoFijo(Index).Text <> "" Then
        
        If Not IsNumeric(txtGastoFijo(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtGastoFijo(Index).Text = ""
            SubSetFocus txtGastoFijo(Index)
        Else
            'sgastfij codigo Descripcion
            txtGastoFijo(Index).Text = Val(txtGastoFijo(Index).Text)
            Sql = DevuelveDesdeBD("Descripcion", "sgastfij", "codigo", txtGastoFijo(Index).Text, "N")
            
            If Sql = "" Then
                MsgBox "No existe el gasto fijo: " & Me.txtGastoFijo(Index).Text, vbExclamation

            End If
        End If
    End If
    Me.txtDescGastoFijo(Index).Text = Sql
     
End Sub











Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
Dim Mal As Boolean
    txtImporte(Index).Text = Trim(txtImporte(Index).Text)
    If txtImporte(Index).Text = "" Then Exit Sub
    Mal = False
    If Not EsNumerico(txtImporte(Index).Text) Then Mal = True

    If Not Mal Then Mal = Not CadenaCurrency(txtImporte(Index).Text, Importe)

    If Mal Then
        txtImporte(Index).Text = ""
        txtImporte(Index).SetFocus

    Else
        txtImporte(Index).Text = Format(Importe, FormatoImporte)
    End If
    




End Sub









'Private Sub PonerNiveles()
'Dim i As Integer
'Dim J As Integer
'
'
'
'
'    Check1(10).Visible = True
'    For i = 1 To vEmpresa.numnivel - 1
'        J = DigitosNivel(i)
'        cad = "Digitos: " & J
'        Check1(i).Visible = True
'        Me.Check1(i).Caption = cad
'
'        'Para los de balance presupuestario
'        Me.ChkCtaPre(i).Visible = True
'        Me.ChkCtaPre(i).Caption = cad
'        'para los de resumen dairio
'        Me.ChkNivelRes(i).Visible = True
'        Me.ChkNivelRes(i).Caption = cad
'
'        'Consolidado
'        Me.ChkConso(i).Visible = True
'        Me.ChkConso(i).Caption = cad
'
'        chkcmp(i).Caption = cad
'        chkcmp(i).Visible = True
'
'        Combo2.AddItem "Nivel :   " & i
'        Combo2.ItemData(Combo2.NewIndex) = J
'    Next i
'    For i = vEmpresa.numnivel To 9
'        Check1(i).Visible = False
'        Me.ChkCtaPre(i).Visible = False
'        Me.ChkNivelRes(i).Visible = False
'        chkcmp(i).Visible = False
'        ChkConso(i).Visible = False
'    Next i
'
'End Sub






Private Sub CargarComboFecha()
'Dim J As Integer
'
'
'QueCombosFechaCargar "0|1|2|"
'
'
''Y ademas deshabilitamos los niveles no utilizados por la aplicacion
'For i = vEmpresa.numnivel To 9
'    Check2(i).Visible = False
'    Me.chkCtaExplo(i).Visible = False
'    chkCtaExploC(i).Visible = False
'    chkAce(i).Visible = False
'Next i
'
'For i = 1 To vEmpresa.numnivel - 1
'    J = DigitosNivel(i)
'    Check2(i).Visible = True
'    Check2(i).Caption = "Digitos: " & J
'    chkCtaExplo(i).Visible = True
'    chkCtaExplo(i).Caption = "Digitos: " & J
'    chkAce(i).Visible = True
'    chkAce(i).Caption = "Digitos: " & J
'    chkCtaExploC(i).Visible = True
'    chkCtaExploC(i).Caption = "Digitos: " & J
'Next i
'
'
'
'
''Cargamos le combo de resalte de fechas
'Combo3.AddItem "Sin remarcar"
'Combo3.ItemData(Combo3.NewIndex) = 1000
'For i = 1 To vEmpresa.numnivel - 1
'    Combo3.AddItem "Nivel " & i
'    Combo3.ItemData(Combo3.NewIndex) = i
'Next i
End Sub




























Private Sub QueCombosFechaCargar(Lista As String)
'Dim L As Integer
'
'L = 1
'Do
'    cad = RecuperaValor(Lista, L)
'    If cad <> "" Then
'        i = Val(cad)
'        With cmbFecha(i)
'            .Clear
'            For Cont = 1 To 12
'                RC = "25/" & Cont & "/2002"
'                RC = Format(RC, "mmmm") 'Devuelve el mes
'                .AddItem RC
'            Next Cont
'        End With
'    End If
'    L = L + 1
'Loop Until cad = ""
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


Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text3(Indice1).Text <> "" And Text3(Indice2).Text <> "" Then
        If CDate(Text3(Indice1).Text) > CDate(Text3(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function





Private Sub txtCtaBanc_GotFocus(Index As Integer)
    PonFoco txtCtaBanc(Index)
End Sub

Private Sub txtCtaBanc_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCtaBanc_LostFocus(Index As Integer)
    txtCtaBanc(Index).Text = Trim(txtCtaBanc(Index).Text)
    If txtCtaBanc(Index).Text = "" Then
        txtDescBanc(Index).Text = ""
        Exit Sub
    End If
    
    cad = txtCtaBanc(Index).Text
    I = CuentaCorrectaUltimoNivelSIN(cad, Sql)
    If I = 0 Then
        MsgBox "NO existe la cuenta: " & txtCtaBanc(Index).Text, vbExclamation
        Sql = ""
        cad = ""
    Else
        cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", cad, "T")
        If cad = "" Then
            MsgBox "Cuenta no asoaciada a ningun banco", vbExclamation
            Sql = ""
            I = 0
        End If
    End If
    
    txtCtaBanc(Index).Text = cad
    Me.txtDescBanc(Index).Text = Sql
    If I = 0 Then PonFoco txtCtaBanc(Index)
    
End Sub

Private Sub txtDias_GotFocus()
    PonFoco txtDias
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDias_LostFocus()
    txtDias.Text = Trim(txtDias.Text)
    If txtDias.Text <> "" Then
        If Not IsNumeric(txtDias.Text) Then
            MsgBox "Numero de dias debe ser numérico", vbExclamation
            txtDias.Text = ""
            SubSetFocus txtDias
        End If
    End If
End Sub



Private Sub txtDpto_GotFocus(Index As Integer)
    PonFoco txtDpto(Index)
End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDpto_LostFocus(Index As Integer)
    
    'Pierde foco
    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    If txtDpto(Index).Text = "" Then
        Me.txtDescDpto(Index).Text = ""
        Exit Sub
    End If
    
    Sql = "NO"
    If txtCta(1).Text = "" Or txtCta(0).Text = "" Then
        MsgBox "Debe seleccionar un unico cliente", vbExclamation
        txtDpto(Index).Text = ""
        Sql = ""
    Else
        If txtCta(1).Text <> txtCta(0).Text Then
            MsgBox "Debe seleccionar un mismo cliente", vbExclamation
            txtDpto(Index).Text = ""
            Sql = ""
        End If
    End If
    
    If Sql <> "" Then
        Sql = ""
        If txtCta(1).Text <> "" Then
            If txtDpto(Index).Text <> "" Then
                If Not IsNumeric(txtDpto(Index).Text) Then
                      MsgBox "Codigo departamento debe ser numerico: " & txtDpto(Index).Text
                      txtDpto(Index).Text = ""
                Else
                      'Comproamos en la BD
                       Set Rs = New ADODB.Recordset
                       cad = "Select descripcion from departamentos where codmacta='" & txtCta(0).Text
                       cad = cad & "' AND Dpto = " & txtDpto(Index).Text
                       Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                       If Not Rs.EOF Then Sql = DBLet(Rs.Fields(0), "T")
                       Rs.Close
                       Set Rs = Nothing
                End If
            End If
        Else
            If txtDpto(Index).Text <> "" Then
                MsgBox "Seleccione un cliente", vbExclamation
                txtDpto(Index).Text = ""
            End If
        End If
    End If
    Me.txtDescDpto(Index).Text = Sql
End Sub

Private Sub txtFPago_GotFocus(Index As Integer)
    PonFoco txtFPago(Index)
End Sub

Private Sub txtFPago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtFPago_LostFocus(Index As Integer)
    If ComprobarCampoENlazado(txtFPago(Index), txtDescFPago(Index), "N") > 0 Then
        If txtFPago(Index).Text <> "" Then
            'Tiene valor.
            Sql = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", txtFPago(Index).Text, "N")
            If Sql = "" Then Sql = "Codigo no encontrado"
            txtDescFPago(Index).Text = Sql
        Else
            'Era un error
            SubSetFocus txtFPago(Index)
        End If
    End If
End Sub




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



Private Function DesdeHasta(Tipo As String, Desde As Integer, Hasta As Integer, Optional Des As String)
Dim C As String
    C = ""
    Select Case Tipo
    Case "F"
        'Son los text3(desde)....
        If Text3(Desde).Text <> "" Then C = C & "Desde " & Text3(Desde).Text
        
        If Text3(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & Text3(Hasta).Text
        End If
        
    Case "C"
        'Cuenta
        If txtCta(Desde).Text <> "" Then C = "Desde " & txtCta(Desde).Text & "-" & DtxtCta(Desde).Text
        
        
        If txtCta(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtCta(Hasta).Text & "-" & DtxtCta(Hasta).Text
        End If
        
        
    Case "FP"
        'FORMA DE PAGO
        If txtFPago(Desde).Text <> "" Then C = "Desde " & txtFPago(Desde).Text & "-" & txtDescFPago(Desde).Text
        
        
        If txtFPago(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtFPago(Hasta).Text & "-" & txtDescFPago(Hasta).Text
        End If
    
    Case "BANCO"
        '    'txtCtaBanc  txtDescBanc
        If txtCtaBanc(Desde).Text <> "" Then C = "Desde " & txtCtaBanc(Desde).Text & "-" & txtDescBanc(Desde).Text
        
        If txtCtaBanc(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtCtaBanc(Hasta).Text & "-" & txtDescBanc(Hasta).Text
        End If
    
    
    Case "S"
        'Serie
        If txtSerie(Desde).Text <> "" Then C = C & "Desde " & txtSerie(Desde).Text
        
        If txtSerie(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtSerie(Hasta).Text
        End If
    
    Case "NF"
        'Numero factura
        If txtNumFac(Desde).Text <> "" Then C = C & "Desde " & txtNumFac(Desde).Text
        
        If txtNumFac(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtNumFac(Hasta).Text
        End If
    
    Case "GF"
        'Gasto fijo
        
        If txtGastoFijo(Desde).Text <> "" Then C = C & "Desde " & txtGastoFijo(Desde).Text & " - " & Me.txtDescGastoFijo(Desde).Text
        
        If txtGastoFijo(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtGastoFijo(Hasta).Text & " - " & Me.txtDescGastoFijo(Hasta).Text
        End If
    
    
    End Select
    If C <> "" Then C = "  " & Des & " " & C
    DesdeHasta = C
End Function


Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
        Me.lblPPAL.Caption = TEXTO
        Me.lbl2.Caption = ""
        Me.ProgressBar1.Value = 0
        Me.FrameProgreso.Visible = True
        Me.Refresh
End Sub





'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

Private Function CobrosPendientesCliente(ByVal ListadoCuentas As String) As Boolean
Dim TieneRemesa As Boolean
Dim RemesaTalones As Boolean
Dim RemesaPagares As Boolean
Dim RemesaEfectos As Boolean
Dim SePuedeRemesar As Boolean
Dim InsertarLinea As Boolean


Dim CadenaInsert As String

    On Error GoTo ECobrosPendientesCliente
    CobrosPendientesCliente = False

    
    'De parametros contapagarepte contatalonpte
    cad = DevuelveDesdeBD("contatalonpte", "paramtesor", "codigo", "1")
    If cad = "" Then cad = "0"
    RemesaTalones = (Val(cad) = 1)
    
    cad = DevuelveDesdeBD("contapagarepte", "paramtesor", "codigo", "1")
    If cad = "" Then cad = "0"
    RemesaPagares = (Val(cad) = 1)
    
    cad = DevuelveDesdeBD("contaefecpte", "paramtesor", "codigo", "1")
    If cad = "" Then cad = "0"
    RemesaEfectos = (Val(cad) = 1)
    

    
    
    
    
    'Trozo basico
    cad = " FROM scobro ,cuentas,sforpa ,stipoformapago"
    cad = cad & " WHERE  scobro.codmacta = cuentas.codmacta"
    cad = cad & " AND scobro.codforpa = sforpa.codforpa"
    cad = cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"

    
    
    'Fecha factura
    RC = CampoABD(Text3(1), "F", "fecfaccl", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(Text3(2), "F", "fecfaccl", False)
    If RC <> "" Then cad = cad & " AND " & RC



    'Se me habia olvidado
    'A G E N T E
    RC = CampoABD(txtAgente(0), "N", "agente", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtAgente(1), "N", "agente", False)
    If RC <> "" Then cad = cad & " AND " & RC




    'Fecha vencimiento
    RC = CampoABD(Text3(19), "F", "fecvenci", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(Text3(20), "F", "fecvenci", False)
    If RC <> "" Then cad = cad & " AND " & RC

    'SERIE
    RC = CampoABD(txtSerie(0), "T", "numserie", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtSerie(1), "T", "numserie", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    'Numero factura
    RC = CampoABD(txtNumFac(0), "T", "codfaccl", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtNumFac(1), "T", "codfaccl", False)
    If RC <> "" Then cad = cad & " AND " & RC
    


    'Cliente
    RC = CampoABD(txtCta(1), "T", "scobro.codmacta", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtCta(0), "T", "scobro.codmacta", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    'Forma PAGO
    RC = CampoABD(txtFPago(0), "T", "scobro.codforpa", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtFPago(1), "T", "scobro.codforpa", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    
    'Cliente con departamento
    'If txtCta(0).Text <> "" Then
    '    If cad <> "" Then cad = cad & " AND "
    '    cad = cad & " scobro.codmacta = '" & txtCta(6).Text & "'"
    'End If
    
    'Departamento
    RC = CampoABD(Me.txtDpto(0), "N", "departamento", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(Me.txtDpto(1), "N", "departamento", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    
    'Solo los NO remesar
    If Me.chkNOremesar.Value = 1 Then
        cad = cad & " AND noremesar = 1 "
    End If
    
    'Docuemtno recibido y devuelto. Los combos  recedocu  Devuelto
    If Me.cboCobro(0).ListIndex > 0 Then cad = cad & " AND recedocu = " & cboCobro(0).ItemData(cboCobro(0).ListIndex)
    If Me.cboCobro(1).ListIndex > 0 Then cad = cad & " AND Devuelto = " & cboCobro(1).ItemData(cboCobro(1).ListIndex)
    
    
    'Y lista de cuentas

    If ListadoCuentas <> "" Then
        NumRegElim = 1
        Sql = ""
        Do
            TotalRegistros = InStr(NumRegElim, ListadoCuentas, "|")
            If TotalRegistros > 0 Then
                If NumRegElim > 1 Then Sql = Sql & ","
                NumRegElim = 2
                Sql = Sql & "'" & Mid(ListadoCuentas, 1, TotalRegistros - 1) & "'"
                ListadoCuentas = Mid(ListadoCuentas, TotalRegistros + 1)
            End If
           
            
        Loop Until TotalRegistros = 0
        NumRegElim = 0
        cad = cad & " AND scobro.codmacta IN (" & Sql & ")"
    End If
    
    
    
    'Si ha marcado alguna forma de pago
    RC = PonerTipoPagoCobro_(True, False)
    If RC <> "" Then cad = cad & " AND tipoformapago IN " & RC
    RC = ""
    
    'Contador
    Sql = "Select count(*) "
    Sql = Sql & cad
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalRegistros = 0
    If Not Rs.EOF Then
        'Total registros
        TotalRegistros = Rs.Fields(0)
    End If
    Rs.Close
    
    If TotalRegistros = 0 Then
        'NO hay registros
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    'Llegados aqui, si k hay registros.
    '100 POR EJEMPLO
    If TotalRegistros > 100 Then
        'Ponemos visible el frame
        MostrarFrame = True
        PonerFrameProgressVisible
    Else
        MostrarFrame = False 'NO lo mostramos
    End If
    
    
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.zpendientes where codusu = " & vUsu.Codigo
    
    'Marzo 2015
    'Si agrupamos por forma de pago, agruparemos antes por Tipo de pago
    If chkFormaPago.Value = 1 Then Conn.Execute "DELETE FROM Usuarios.zsimulainm where codusu = " & vUsu.Codigo
    
    
    
    
    
    Sql = "SELECT scobro.*, cuentas.nommacta, nifdatos,stipoformapago.descformapago ,stipoformapago.tipoformapago,nomforpa " & cad
    ' ----------------
    '  20 Diciembre 2005
    '  Remesados o no remesados
    '
    CONT = 1
    If Me.ChkAgruparSituacion.Value = 1 Then
        '
        CONT = 0
    Else
        If Me.chkEfectosContabilizados.Value = 1 Then CONT = 0
    End If
    If CONT = 1 Then Sql = Sql & " AND codrem is null"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    TieneRemesa = False
    Sql = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,"
    Sql = Sql & "codforpa, nomforpa, codmacta,nombre, fecVto, importe, pag_cob,gastos,vencido,Situacion"
    'Nuevo Enero 2009
    'Si esta apaisado ponemos los departamentos
    If Me.chkApaisado(0).Value = 1 Then
        Sql = Sql & ",coddirec,nomdirec"
    Else
        'Metemos el NIF para futors listados. Pej. El de cobors por cliente lo pondra
        Sql = Sql & ",nomdirec"
    End If
    Sql = Sql & ",devuelto,recibido"
    'SQL = SQL & ",observa) VALUES (" & vUsu.Codigo & ",'"
    'Dic 2013 . Acelerar proceso
    Sql = Sql & ",observa) VALUES "
    
    
    CadenaInsert = "" 'acelerar carga datos
    Fecha = CDate(Text3(0).Text)
    While Not Rs.EOF
        CONT = CONT + 1
        If MostrarFrame Then
            lbl2.Caption = "Registro: " & CONT
            lbl2.Refresh
        End If
        
        
        'If Rs!codmacta = "4300019" Then Stop
        
        cad = Rs!NUmSerie & "','" & Format(Rs!codfaccl, "0000") & "','" & Format(Rs!fecfaccl, FormatoFecha) & "'," & Rs!numorden
        
        'Modificacion. Enero 2010. Tiene k aparacer la forma de pago, no el tipo
        'Cad = Cad & "," & Rs!codforpa & ",'" & DevNombreSQL(Rs!descformapago) & "','"
        cad = cad & "," & Rs!codforpa & ",'" & DevNombreSQL(Rs!nomforpa) & "','"
        
        cad = cad & Rs!codmacta & "','" & DevNombreSQL(Rs!Nommacta) & "','"
        cad = cad & Format(Rs!FecVenci, FormatoFecha) & "',"
        cad = cad & TransformaComasPuntos(CStr(Rs!ImpVenci)) & ","
        'Cobrado, si no es nulo
        If Not IsNull(Rs!impcobro) Then
            cad = cad & TransformaComasPuntos(CStr(Rs!impcobro))
        Else
            cad = cad & "0"
        End If
        
        'Gastos
        cad = cad & "," & TransformaComasPuntos(DBLet(Rs!Gastos, "N"))
        
        If Fecha > Rs!FecVenci Then
            cad = cad & ",1"
        Else
            cad = cad & ",0"
        End If

        'Hay que añadir la situacion. Bien sea juridica....
        ' Si NO agrupa por situacion, en ese campo metere la referencia del cobro (rs!referencia)
         'vbTalon = 2 vbPagare = 3
        InsertarLinea = True
        
        If Me.ChkAgruparSituacion.Value = 0 Then
            cad = cad & ",'" & DevNombreSQL(DBLet(Rs!referencia, "T")) & "'"
            
            'Si no agrupa por situacion de vto y no tiene el riesgo marcado
            'Talones pagares
            'Si se ha recepcionado documento, NO deben salir
            
            
            'Enero 2010
            'Comentamos esto ya que tiene la marca de recibidos si/no
'            If Me.chkEfectosContabilizados.Value = 0 Then
'                If Val(Rs!tipoformapago) = vbTalon Or Val(Rs!tipoformapago) = vbPagare Then
'                    If DBLet(Rs!recedocu, "N") = 1 Then InsertarLinea = False
'                End If
'            End If

            
        Else
            'La situacion.
            'Si esta en situacion juridica.
            ' Si no, si esta devuelto y no es una remesa
            ' y luego si es una remesa, sitaucion o devuelto
            If Rs!situacionjuri = 1 Then
                cad = cad & ",'SITUACION JURIDICA'"
            Else
                'Cambio Marzo 2009
                ' Ahora tb se remesan los pagares y talones
                
                If Not IsNull(Rs!siturem) Then
                    TieneRemesa = True
                    cad = cad & ",'R" & Format(Rs!AnyoRem, "0000") & Format(Rs!CodRem, "0000000000") & "'"
                    
                Else
                    
                    If Rs!Devuelto = 1 Then
                        cad = cad & ",'DEVUELTO'"
                    Else
                            
                        SePuedeRemesar = False
                        If RemesaEfectos Then SePuedeRemesar = Rs!tipoformapago = vbTipoPagoRemesa
                        If RemesaPagares Then SePuedeRemesar = Rs!tipoformapago = vbPagare
                        If RemesaTalones Then SePuedeRemesar = Rs!tipoformapago = vbTalon
                        
                    
                        If Not SePuedeRemesar Then
                            cad = cad & ",'PENDIENTE COBRO'"
                        Else
                            cad = cad & ",'PENDIENTE REMESAR'" '& Rs!anyorem
                        End If
                        
                        
                        
                        'Talones pagares
                        'Si se ha recepcionado documento, NO deben salir
                        'ENERO 2010
                        'Tiene la marca de recibidos
                        
                        'If Val(Rs!tipoformapago) = vbTalon Or Val(Rs!tipoformapago) = vbPagare Then
                        '    If Me.chkEfectosContabilizados.Value = 0 Then
                        '        If DBLet(Rs!recedocu, "N") = 1 Then InsertarLinea = False
                        '    End If
                        'End If
                        
                    
                    End If  'De devuelto
               End If  'de SITUREM null
            End If ' de situacion juridica
        End If  'de agrupa por sitacuib
        cad = cad & ","
        If Me.chkApaisado(0).Value = 1 Then
            'SI carga departamentos. Esto podriamos mejorar la velocidad si
            'pregarmos un rs o en la select linkamos con departamento...
            If IsNull(Rs!departamento) Then
                cad = cad & "NULL,NULL,"
            Else
                cad = cad & "'" & Rs!departamento & "','"
                cad = cad & DevNombreSQL(DevuelveDesdeBD("Descripcion", "departamentos", "codmacta = '" & Rs!codmacta & "' AND dpto", Rs!departamento, "N")) & "',"
            End If
            
        Else
            'Nif datos
            'Stop
             cad = cad & "'" & DevNombreSQL(DBLet(Rs!nifdatos, "T")) & "',"
        End If
        
        If DBLet(Rs!Devuelto, "N") = 0 Then
            cad = cad & "'',"
        Else
            cad = cad & "'S',"
        End If
        If DBLet(Rs!recedocu, "N") = 0 Then
            cad = cad & "''"
        Else
            cad = cad & "'S'"
        End If
            
        cad = cad & ",'"
        If Me.ChkObserva.Value Then
            cad = cad & DevNombreSQL(DBLet(Rs!Obs, "T"))
'        Else
'            Cad = Cad & "''"
        End If
        cad = cad & "')"
        
        If InsertarLinea Then
        
            CadenaInsert = CadenaInsert & ", (" & vUsu.Codigo & ",'" & cad
        
            If Len(CadenaInsert) > 20000 Then
                cad = Sql & Mid(CadenaInsert, 2)
                Conn.Execute cad
                CadenaInsert = ""
            End If
            'Cad = SQL & Cad
            'Conn.Execute Cad
        Else
            'Tiramos para atras el contador total
            CONT = CONT - 1
        End If
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    
    If Len(CadenaInsert) > 0 Then
        cad = Sql & Mid(CadenaInsert, 2)
        Conn.Execute cad
        CadenaInsert = ""
    End If

    
    'Si esta seleccacona SITIACUIN VENCIMIENTO
    ' y tenia remesas , entonces updateo la tabla poniendo
    ' la situacion de la remesa
    If TieneRemesa Then
        cad = "Select codigo,anyo,  descsituacion"
        cad = cad & " from remesas left join tiposituacionrem on situacion=situacio"
        Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Debug.Print Rs!Codigo
            If Not IsNull(Rs!descsituacion) Then
                cad = "R" & Format(Rs!Anyo, "0000") & Format(Rs!Codigo, "0000000000")
                cad = " WHERE situacion='" & cad & "'"
                cad = "UPDATE Usuarios.zpendientes set Situacion='Remesados: " & Rs!descsituacion & "' " & cad
                Conn.Execute cad
            End If
            Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    'Marzo 2015.
    'Nivel de anidacion para los agrupados por forma de pago
    ' que es TIPO DE PAGO
    If chkFormaPago.Value = 1 Then
    
        cad = "select codforpa from Usuarios.zpendientes where codusu =" & vUsu.Codigo & " group by 1"
        Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        While Not Rs.EOF
            cad = cad & ", " & Rs!codforpa
            Rs.MoveNext
        Wend
        Rs.Close
        
        If cad <> "" Then
            cad = Mid(cad, 2)
            cad = " and codforpa IN (" & cad & ")"
            cad = " FROM sforpa , stipoformapago WHERE sforpa.tipforpa=stipoformapago.tipoformapago" & cad
            cad = "SELECT " & vUsu.Codigo & ",codforpa,tipoformapago,descformapago " & cad
            cad = "INSERT INTO Usuarios.zsimulainm(codusu,codigo,conconam,nomconam) " & cad
        
            Conn.Execute cad
        End If
    End If
    
    
    
    'Voy a comprobar si ha metido algun registo
    espera 0.3
    Sql = "Select count(*) FROM  Usuarios.zpendientes where codusu = " & vUsu.Codigo
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not Rs.EOF Then CONT = DBLet(Rs.Fields(0), "N")
    Rs.Close
    If CONT = 0 Then
        MsgBox "No se ha generado ningun valor(2)", vbExclamation
    Else
        CobrosPendientesCliente = True 'Para imprimir
    End If
    Exit Function
ECobrosPendientesCliente:
    MuestraError Err.Number, Err.Description
End Function



Private Function PagosPendienteProv(ByVal ListadoCuentas As String) As Boolean
'Dim MismaClavePrimaria As String


    On Error GoTo EPagosPendienteProv
    PagosPendienteProv = False
    
    'Trozo basico de PAGOS
    cad = "FROM spagop ,cuentas ,sforpa,stipoformapago"
    cad = cad & " WHERE spagop.ctaprove = cuentas.codmacta"
    cad = cad & " AND spagop.codforpa = sforpa.codforpa"
    cad = cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"

    
    'Fecha
    RC = CampoABD(Text3(3), "F", "fecefect", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(Text3(4), "F", "fecefect", False)
    If RC <> "" Then cad = cad & " AND " & RC

    'Cliente
    RC = CampoABD(txtCta(2), "T", "ctaprove", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtCta(3), "T", "ctaprove", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    
    'FORMA PAGO
    RC = CampoABD(txtFPago(6), "N", "spagop.codforpa", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtFPago(7), "N", "spagop.codforpa", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    
    
    
    
    
    'Y lista de cuentas

    If ListadoCuentas <> "" Then
        NumRegElim = 1
        Sql = ""
        Do
            TotalRegistros = InStr(NumRegElim, ListadoCuentas, "|")
            If TotalRegistros > 0 Then
                If NumRegElim > 1 Then Sql = Sql & ","
                NumRegElim = 2
                Sql = Sql & "'" & Mid(ListadoCuentas, 1, TotalRegistros - 1) & "'"
                ListadoCuentas = Mid(ListadoCuentas, TotalRegistros + 1)
            End If
           
            
        Loop Until TotalRegistros = 0
        NumRegElim = 0
        cad = cad & " AND spagop.ctaprove IN (" & Sql & ")"
        
    End If
    
    
    'ORDEN
    cad = cad & " ORDER BY numfactu"
   
    
    
    
    'Contador
    Sql = "Select count(*) "
    Sql = Sql & cad
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalRegistros = 0
    If Not Rs.EOF Then
        'Total registros
        TotalRegistros = Rs.Fields(0)
    End If
    Rs.Close
    
    If TotalRegistros = 0 Then
        'NO hay registros
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    'Llegados aqui, si k hay registros.
    '100 POR EJEMPLO
    If TotalRegistros > 100 Then
        'Ponemos visible el frame
        MostrarFrame = True
        PonerFrameProgressVisible
    Else
        MostrarFrame = False 'NO lo mostramos
    End If
    
    
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.zpendientes where codusu = " & vUsu.Codigo
    
    Sql = "SELECT spagop.*, cuentas.nommacta, stipoformapago.descformapago, stipoformapago.siglas,nomforpa " & cad
    
    'Cambiamos
''''''    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''''    'Compruebo si hay repetidos en fecfactu|numfactu|siglas
''''''    SQL = ""
''''''    MismaClavePrimaria = "|"
''''''    While Not RS.EOF
''''''        RC = Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu
''''''        If RC = SQL Then
''''''            RC = RC & "|"
''''''            If InStr(1, MismaClavePrimaria, "|" & RC) = 0 Then MismaClavePrimaria = MismaClavePrimaria & RC
''''''        Else
''''''            SQL = RC
''''''        End If
''''''        RS.MoveNext
''''''    Wend
''''''    SQL = RS.Source
''''''    RS.Close
    
    'Vuelvo a abrirlo
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Agosto 2013
    'Añadimos en campo SITUACION donde pondra si esta emitido o no (emitdocum)
    
    'Mayo 2014
    'La factura la metemos en nomdirec. Asi NO da error duplicados
    
    CONT = 0
    Sql = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,nomdirec,"
    Sql = Sql & "codforpa, nomforpa, codmacta,nombre, fecVto, importe, pag_cob,vencido,situacion) VALUES (" & vUsu.Codigo & ",'"
    Fecha = CDate(Text3(5).Text)
    DevfrmCCtas = ""
    While Not Rs.EOF
        CONT = CONT + 1
        If MostrarFrame Then
            lbl2.Caption = "Registro: " & CONT
            lbl2.Refresh
        End If
        
'        'Por si se repiten misma fecfactura, mismo numero factura, mismo tipo de pago
'        If MismaClavePrimaria <> "" Then
'            'Hay claves repetidas no tiene en cuenta el vto
'            RC = "|" & Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu & "|"
'            'Enero 2011
'            RC = "|" & Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu & "#" & RS!numorden & "|"
'
'
'            If InStr(1, MismaClavePrimaria, RC) > 0 Then
'                RC = DevNombreSQL(RS!numfactu)
'                RC = FijaNumeroFacturaRepetido(RC)
'                Cad = RS!siglas & "','" & RC & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'            Else
'                'El mismo de abajo
'                Cad = RS!siglas & "','" & DevNombreSQL(RS!numfactu) & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'            End If
'        Else
'            'El procedimiento de antes
'            Cad = RS!siglas & "','" & DevNombreSQL(RS!numfactu) & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'        End If
'
        
        'mayo 2014
        cad = Rs!siglas & "','" & Format(CONT, "00000") & "','" & Format(Rs!FecFactu, FormatoFecha) & "'," & Rs!numorden & ",'" & DevNombreSQL(Rs!NumFactu) & "'"
        
        
        'optMostraFP
        cad = cad & "," & Rs!codforpa & ",'"
        If Me.optMostraFP(0).Value Then
            cad = cad & DevNombreSQL(Rs!descformapago)
        Else
            cad = cad & DevNombreSQL(Rs!nomforpa)
        End If
        cad = cad & "','" & Rs!ctaprove & "','" & DevNombreSQL(Rs!Nommacta) & "','"
        cad = cad & Format(Rs!fecefect, FormatoFecha) & "',"
        cad = cad & TransformaComasPuntos(CStr(Rs!ImpEfect)) & ","
        'Cobrado, si no es nulo
        If Not IsNull(Rs!imppagad) Then
            cad = cad & TransformaComasPuntos(CStr(Rs!imppagad))
        Else
            cad = cad & "0"
        End If
        If Fecha > Rs!fecefect Then
            cad = cad & ",1"
        Else
            cad = cad & ",0"
        End If
        
        'Agosto 2013
        'Si esta en un tal-pag
        cad = cad & ",'"
        If DBLet(Rs!emitdocum, "N") > 0 Then cad = cad & "*"
        
        cad = cad & "')"  'lleva el apostrofe
        cad = Sql & cad
        Conn.Execute cad
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
     
    PagosPendienteProv = True 'Para imprimir
    Exit Function
EPagosPendienteProv:
    MuestraError Err.Number, Err.Description
End Function



Private Function FijaNumeroFacturaRepetido(Numerofactura) As String
Dim I As Integer
Dim Aux As String
        If Len(Numerofactura) >= 10 Then
            MsgBox "Clave duplicada. Imposible insertar. " & Rs!NumFactu & ": " & Rs!FecFactu, vbExclamation
            FijaNumeroFacturaRepetido = Numerofactura
            Exit Function
        End If
        
        'Añadiremos guienos por detras
        For I = Len(Numerofactura) To 10
            'Añadirenos espacios en blanco al final
            Aux = Rs!NumFactu & String(I - Len(Numerofactura), "_")
            If InStr(1, DevfrmCCtas, "|" & Aux & "|") = 0 Then
                'Devolvemos este numero de factura
                FijaNumeroFacturaRepetido = Aux
                If DevfrmCCtas = "" Then DevfrmCCtas = "|"
                DevfrmCCtas = DevfrmCCtas & Aux & "|"
                Exit Function
            End If
            
        Next I
        
        'Si llega aqui probaremos con los -
        For I = Len(Numerofactura) + 1 To 10
            'Añadirenos espacios en blanco al final
            Aux = String(I - Len(Numerofactura), "_") & Rs!NumFactu
            If InStr(1, DevfrmCCtas, "|" & Aux & "|") = 0 Then
                'Devolvemos este numero de factura
                FijaNumeroFacturaRepetido = Aux
                DevfrmCCtas = DevfrmCCtas & Aux & "|"
                Exit Function
            End If
            
        Next I
End Function


Private Sub txtNumero_GotFocus(Index As Integer)
    PonFoco txtNumero(Index)
End Sub



Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtNumFac_GotFocus(Index As Integer)
    PonFoco txtNumFac(Index)
End Sub

Private Sub txtNumFac_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtNumFac_LostFocus(Index As Integer)
    txtNumFac(Index).Text = Trim(txtNumFac(Index).Text)
    If txtNumFac(Index).Text = "" Then Exit Sub
    
    If Not IsNumeric(txtNumFac(Index).Text) Then
        MsgBox "Campo debe ser numerico.", vbExclamation
        txtNumFac(Index).Text = ""
        PonFoco txtNumFac(Index)
    End If
End Sub

Private Sub txtRem_GotFocus(Index As Integer)
    PonFoco txtRem(Index)
End Sub

Private Sub txtRem_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtSerie_GotFocus(Index As Integer)
    PonFoco txtSerie(Index)
End Sub

Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtSerie_LostFocus(Index As Integer)
    txtSerie(Index).Text = UCase(txtSerie(Index))
End Sub

Private Sub txtVarios_GotFocus(Index As Integer)
    PonFoco txtVarios(Index)
End Sub

Private Sub txtVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Function ListadoRemesas() As Boolean
Dim Aux As String
    On Error GoTo EListadoRemesas
    ListadoRemesas = False
    
    Sql = ""
    RC = CampoABD(txtRem(0), "N", "codigo", True)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    RC = CampoABD(txtRem(1), "N", "codigo", False)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    RC = CampoABD(txtRem(2), "N", "anyo", True)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    RC = CampoABD(txtRem(3), "N", "anyo", False)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    'Tipo remesa
    RC = RemesaSeleccionTipoRemesa(chkTipoRemesa(0).Value = 1, chkTipoRemesa(1).Value = 1, chkTipoRemesa(2).Value = 1)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    Set Rs = New ADODB.Recordset
    
    'ANTES
    RC = "SELECT remesas.*,nommacta from remesas,cuentas "
    RC = RC & " WHERE remesas.codmacta = cuentas.codmacta"
    If Sql <> "" Then RC = RC & " AND " & Sql
    
    
    'AHORA
    RC = "Select codigo,anyo, fecremesa,tiporemesa.descripcion as d1,descsituacion,remesas.codmacta,nommacta,"
    RC = RC & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    RC = RC & " from cuentas,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    RC = RC & " and situacio=situacion"
    If Sql <> "" Then RC = RC & " AND " & Sql
    
    
    Rs.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ninguna remesa para listar", vbExclamation
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    
    Sql = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc                tipo   situacion
    Sql = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4, texto5,importe1,  fecha1) VALUES ("
    'Para insertar en la otra
    cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe,haber) VALUES ("
    
    
    
    TotalRegistros = 0
    
    While Not Rs.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = Rs!Anyo * 100000 + Rs!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(Rs!Descripcion, "T")) & "','" & DevNombreSQL(Rs!Nommacta) & "',"
        'TIPO   situacion
        
        RC = RC & "'" & DBLet(Rs!d1, "t") & "','" & DBLet(Rs!descsituacion, "T") & "','"
        
        'Tipo remesa
        If Rs!Tiporem = 2 Then
            RC = RC & "PAG"
        ElseIf Rs!Tiporem = 3 Then
            RC = RC & "TAL"
        Else
            RC = RC & "EFE"
        End If
        RC = RC & "'," & TransformaComasPuntos(CStr(Rs!Importe)) & ",'" & Format(Rs!fecremesa, FormatoFecha) & "')"
    
        RC = Sql & RC
        Conn.Execute RC
       
        I = 1
        If Me.chkRem(0).Value = 1 Then
            'fecfaccl,scobro.codmacta,siturem,impcobro,impvenci,gastos,codfaccl,numserie,codbanco,codsucur,digcontr,cuentaba
            RC = "SELECT fecfaccl,scobro.codmacta,siturem,impcobro,impvenci,gastos,codfaccl,numserie,codbanco,codsucur,digcontr,scobro.cuentaba,nommacta"
            RC = RC & " ,fecvenci,scobro.iban from scobro,cuentas where codrem=" & Rs!Codigo & " AND anyorem =" & Rs!Anyo
            RC = RC & " AND cuentas.codmacta = scobro.codmacta  ORDER BY "
            If Me.optOrdenRem(1).Value Then
                'Codmacta
                RC = RC & "scobro.codmacta,numserie,codfaccl,fecfaccl"
            ElseIf Me.optOrdenRem(2).Value Then
                'Nommacta
                RC = RC & "nommacta,numserie,codfaccl,fecfaccl"
            ElseIf Me.optOrdenRem(0).Value Then
                'Numero factura
                RC = RC & "numserie,codfaccl,fecfaccl"
            Else
                'fcto
                RC = RC & "fecvenci,numserie,codfaccl,fecfaccl"
            
            End If
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'If CONT = 200900004 Then Stop
                'cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien,
                'fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe
                RC = vUsu.Codigo & "," & CONT & ",'" & Format(miRsAux!fecfaccl, FormatoFecha) & "',"
                RC = RC & I & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                'Importe = miRsAux!impvenci - DBLet(miRsAux!impcobro, "N") + DBLet(miRsAux!Gastos, "N")
                If miRsAux!siturem > "B" Then
                    'No deberia ser NULL
                    Importe = DBLet(miRsAux!impcobro, "N")
                Else
                    Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
                End If
                RC = RC & Format(miRsAux!codfaccl, "000000000") & "','"
                
                'Aqui pondre el CCC para los efectos
                '---------------------------------------------------
                'rc=rc & codbanco,codsucur,digcontr,scobro.cuentaba
                Aux = ""
                If Rs!Tiporem = 1 Then
                        If IsNull(miRsAux!codbanco) Then
                            Aux = "0000"
                        Else
                            Aux = Format(miRsAux!codbanco, "0000")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!codsucur) Then
                            Aux = Aux & "0000"
                        Else
                            Aux = Aux & Format(miRsAux!codsucur, "0000")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!digcontr) Then
                            Aux = Aux & "**"
                        Else
                            Aux = Aux & Format(miRsAux!digcontr, "00")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!Cuentaba) Then
                            Aux = Aux & "0000"
                        Else
                            Aux = Aux & Format(miRsAux!Cuentaba, "0000000000")
                        End If
                Else
                    'Talon / Pagare. Si tiene numero puesto lo pondre
                 
                End If
                
                'Nuevo ENERO 2010
                'Fecha vto
                Aux = DBLet(miRsAux!IBAN, "T") & Aux
                If Len(Aux) > 24 Then Aux = Mid(Aux, 1, 24)
                Aux = Aux & "|" & Format(miRsAux!FecVenci, "dd/mm/yy")
                
                RC = RC & Aux & "'," & TransformaComasPuntos(CStr(Importe))
                
                'En el haber pongo el ascii de la serie
                '--------------------------------------
                RC = RC & "," & Asc(Left(DBLet(miRsAux!NUmSerie, "T") & " ", 1)) & ")"
                RC = cad & RC
            
                Conn.Execute RC
            
                'Sig
                I = I + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
       
        Else
            'Tenemos k insertar una unica linea a blancos
            RC = vUsu.Codigo & "," & CONT & ",'1999-12-31'," & I & ",'','','','',0,0)"
            RC = cad & RC
            
            Conn.Execute RC
        End If
        TotalRegistros = TotalRegistros + 1
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    
    
    
    Set Rs = Nothing
    Set miRsAux = Nothing
    
    If Me.chkRem(0).Value = 1 Then
        If TotalRegistros = 0 Then
            MsgBox "No hay vencimientos asociados a las remesas", vbExclamation
            Exit Function
        End If
    End If
    ListadoRemesas = True
    Exit Function
EListadoRemesas:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    Set miRsAux = Nothing

End Function









Private Function ListadoRemesasBanco() As Boolean
Dim Aux As String
Dim Cad2 As String
Dim J As Integer
    On Error GoTo EListadoRemesas
    ListadoRemesasBanco = False
    
    Sql = ""
    RC = CampoABD(txtRem(0), "N", "codigo", True)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    RC = CampoABD(txtRem(1), "N", "codigo", False)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    RC = CampoABD(txtRem(2), "N", "anyo", True)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    RC = CampoABD(txtRem(3), "N", "anyo", False)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    'Tipo remesa
    RC = RemesaSeleccionTipoRemesa(chkTipoRemesa(0).Value = 1, chkTipoRemesa(1).Value = 1, chkTipoRemesa(2).Value = 1)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    Set Rs = New ADODB.Recordset
    
    'ANTES
    RC = "SELECT remesas.*,nommacta from remesas,cuentas "
    RC = RC & " WHERE remesas.codmacta = cuentas.codmacta"
    If Sql <> "" Then RC = RC & " AND " & Sql
    
    
    'AHORA
    RC = "Select codigo,anyo, fecremesa,tiporemesa.descripcion as d1,descsituacion,remesas.codmacta,nommacta,"
    RC = RC & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    RC = RC & " from cuentas,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    RC = RC & " and situacio=situacion"
    If Sql <> "" Then RC = RC & " AND " & Sql
    
    
    Rs.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ninguna remesa para listar", vbExclamation
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    
    Sql = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc                tipo   situacion
    Sql = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4, texto5,importe1,  fecha1,observa1) VALUES ("
    
    
    
    TotalRegistros = 0
    
    While Not Rs.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = Rs!Anyo * 100000 + Rs!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(Rs!Descripcion, "T")) & "','" & DevNombreSQL(Rs!Nommacta) & "',"
        'TIPO   situacion
        
        RC = RC & "'" & DBLet(Rs!d1, "t") & "','" & DBLet(Rs!descsituacion, "T") & "','"
        
        'Tipo remesa
        If Rs!Tiporem = 2 Then
            RC = RC & "PAG"
        ElseIf Rs!Tiporem = 3 Then
            RC = RC & "TAL"
        Else
            RC = RC & "EFE"
        End If
        RC = RC & "'," & TransformaComasPuntos(CStr(Rs!Importe)) & ",'" & Format(Rs!fecremesa, FormatoFecha) & "','"
        
        Cad2 = "Select * from ctabancaria where codmacta = '" & Rs!codmacta & "'"
        miRsAux.Open Cad2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad2 = "NO ENCONTRADO"
        If Not miRsAux.EOF Then
            Cad2 = Trim(DBLet(miRsAux!IBAN, "T") & " ") & Format(DBLet(miRsAux!Entidad, "N"), "0000") & " " & Format(DBLet(miRsAux!Oficina, "N"), "0000") & " "
            If IsNull(miRsAux!Control) Then
                Cad2 = Cad2 & "*"
            Else
                Cad2 = Cad2 & miRsAux!Control
            End If
            Cad2 = Cad2 & " " & Format(DBLet(miRsAux!CtaBanco, "N"), "0000000000")
        End If
        miRsAux.Close
        RC = RC & Cad2 & "')"
        'ctabancaria(entidad,oficina,control,ctabanco)
        Cad2 = ""
        
        RC = Sql & RC
        Conn.Execute RC
       
        I = 1
        
            'Voy a comprobar que existen
            RC = "SELECT codmacta,reftalonpag FROM scobro "
            RC = RC & "  WHERE codrem=" & Rs!Codigo & " AND anyorem =" & Rs!Anyo
            RC = RC & " GROUP BY 1,2 "
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad2 = ""
            While Not miRsAux.EOF
                Cad2 = Cad2 & " scarecepdoc.codmacta = '" & miRsAux!codmacta & "' AND numeroref = '" & DevNombreSQL(miRsAux!reftalonpag) & "'|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If Cad2 = "" Then
                MsgBox "Error obteniendo cuenta/referenciatalon", vbExclamation
                Rs.Close
                Exit Function
            End If
                
            'Comprobare que existen y de paso los inserto
            While Cad2 <> ""
                J = InStr(1, Cad2, "|")
                Aux = Mid(Cad2, 1, J - 1)
                Cad2 = Mid(Cad2, J + 1)
                
                'RC = "SELECT * FROM scarecepdoc ,slirecepdoc,cuentas WHERE codigo=id AND scarecepdoc.codmacta=cuentas.codmacta AND " & Aux
                RC = "SELECT * FROM scarecepdoc ,cuentas WHERE  scarecepdoc.codmacta=cuentas.codmacta AND " & Aux
               
                miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    MsgBox "No se encuentra la referencia; " & Aux, vbExclamation
                    miRsAux.Close
                    Rs.Close
                    Exit Function
                End If
                
                While Not miRsAux.EOF
            
                
                
                
                
                    
                    'Para insertar en la otra
                    cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu,  nommacta,codmacta, numdocum, ampconce, debe,haber) VALUES ("
                
                    'Trampas:  Entre codmacta numdocum llevare el numero de talon. Ya que suman 20 y reftal es len20
                    RC = vUsu.Codigo & "," & CONT & ",'" & Format(miRsAux!fechavto, FormatoFecha) & "',"
                    RC = RC & I & ",'" & DevNombreSQL(miRsAux!Nommacta) & "','"
                    Importe = DBLet(miRsAux!Importe, "N")
                    
                    'Referencia talon
                    Aux = DevNombreSQL(miRsAux!numeroref)
                    If Len(Aux) > 10 Then
                        RC = RC & Mid(Aux, 1, 10) & "','" & Mid(Aux, 11)
                    Else
                        RC = RC & Aux & "','"
                    End If
                    'Banco
                    RC = RC & "','" & DevNombreSQL(miRsAux!banco) & "',"
                    
                    'Talon / Pagare. Si tiene numero puesto lo pondre
                    RC = RC & TransformaComasPuntos(CStr(Importe))
                    
                    'En el haber pongo el ascii de la serie
                    '--------------------------------------
                    RC = RC & ",0)"
                    RC = cad & RC
                
                    Conn.Execute RC
                
                    'Sig
                    I = I + 1
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
            Wend

        TotalRegistros = TotalRegistros + 1
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    
    
    
      Set Rs = Nothing
    Set miRsAux = Nothing
    
    If Me.chkRem(0).Value = 1 Then
        If TotalRegistros = 0 Then
            MsgBox "No hay vencimientos asociados a las remesas", vbExclamation
            Exit Function
        End If
    End If
    ListadoRemesasBanco = True
    Exit Function
EListadoRemesas:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    Set miRsAux = Nothing

End Function




'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'               CREDITO CAUCION
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Private Function ListadoTransferencias() As Boolean
Dim Importe As Currency

    On Error GoTo EListadoTransferencias
    ListadoTransferencias = False
    
    Sql = ""
    RC = CampoABD(txtNumero(0), "N", "codigo", True)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    RC = CampoABD(txtNumero(1), "N", "codigo", False)
    If RC <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & RC
    End If
    
    
    cad = RC
    
    Set Rs = New ADODB.Recordset
    
    RC = "SELECT stransfer.*,nommacta from stransfer"
    If Opcion = 13 Then RC = RC & "cob"
    RC = RC & " as stransfer,cuentas "
    RC = RC & " WHERE stransfer.codmacta = cuentas.codmacta"
    If Sql <> "" Then RC = RC & " AND " & Sql
    
    Rs.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ninguna valor para listar", vbExclamation
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    
    Sql = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    If Opcion = 13 Then Conn.Execute "Delete from usuarios.zcuentas where codusu =" & vUsu.Codigo
        
    
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc
    Sql = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, importe1,  fecha1) VALUES ("
    'Para insertar en la otra
    cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe) VALUES ("
    
    
    

    
    While Not Rs.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = Rs!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(Rs!Descripcion, "T")) & "','" & DevNombreSQL(Rs!Nommacta) & "',"
        RC = RC & TransformaComasPuntos("0") & ",'" & Format(Rs!Fecha, FormatoFecha) & "')"
    
        RC = Sql & RC
        Conn.Execute RC
       
        I = 1
     
            
            If Opcion = 13 Then
                RC = "scobro"
            Else
                RC = "spagop"
            End If
            RC = "SELECT " & RC & ".*,nommacta from cuentas," & RC
            RC = RC & " WHERE transfer = " & Rs!Codigo
            RC = RC & " AND cuentas.codmacta = "
            If Opcion = 13 Then
                RC = RC & " scobro.codmacta "
                RC = RC & " ORDER BY scobro.codmacta,fecfaccl"
            Else
                RC = RC & " spagop.ctaprove "
                RC = RC & " ORDER BY ctaprove,fecfactu"
            End If
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien,
                'fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe "
                If Opcion = 13 Then
                    Fecha = miRsAux!fecfaccl
                Else
                    Fecha = miRsAux!FecFactu
                End If
                RC = vUsu.Codigo & "," & CONT & ",'" & Format(Fecha, FormatoFecha) & "',"
                RC = RC & I & ",'"
                If Opcion = 13 Then
                    RC = RC & miRsAux!codmacta
                Else
                    RC = RC & miRsAux!ctaprove
                End If
                
                RC = RC & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                
                
                'Cuenta
                If Opcion <> 13 Then
                    RC = RC & DevNombreSQL(miRsAux!NumFactu) & "','"
                    
                    'Noviembre 2013
                    'Añadimos el IBAN
                    
                    RC = RC & Trim(DBLet(miRsAux!IBAN, "T") & " " & Format(DBLet(miRsAux!Entidad, "T"), "0000")) & " " & Format(DBLet(miRsAux!Oficina, "T"), "0000") & " "
                    RC = RC & Mid(DBLet(miRsAux!CC, "T") & "**", 1, 2) & " " & Right(String(10, "0") & DBLet(miRsAux!Cuentaba, "T"), 10)
                    Importe = miRsAux!ImpEfect - (DBLet(miRsAux!imppagad, "N"))
                    RC = RC & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
                Else
                    RC = RC & DevNombreSQL(miRsAux!codfaccl) & "','"
                    
                    CadenaDesdeOtroForm = "NO"
                    If DBLet(miRsAux!codbanco, "N") > 0 Then
                        'Es especifico para ESCALONO, pero no molesta a nadie
                        If DBLet(miRsAux!Cuentaba, "T") = "8888888888" Then
                            'Seguira poniendo  cuenta no domiciliada
                        Else
                            CadenaDesdeOtroForm = ""
                        End If
                    End If
                    If CadenaDesdeOtroForm = "" Then
                        'OK, ponemos la cuenta
                        CadenaDesdeOtroForm = Trim(DBLet(miRsAux!IBAN, "T") & " " & Format(DBLet(miRsAux!codbanco, "N"), "0000")) & " " & Format(DBLet(miRsAux!codsucur, "N"), "0000") & " "
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "**  ******" & Right(String(4, "0") & DBLet(miRsAux!Cuentaba, "T"), 4)
                        
                    Else
                        'CUENTANODOMICILIADA
                        CadenaDesdeOtroForm = "NODOM"  'en el rpt haremos un replace
                    End If
                    RC = RC & CadenaDesdeOtroForm
                    Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
                    RC = RC & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
                End If
                RC = cad & RC
            
                Conn.Execute RC
            
                'Sig
                I = I + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
       
'        Else
'            'Tenemos k insertar una unica linea a blancos
'            RC = vUsu.Codigo & "," & CONT & ",''," & I & ",'','','','',0)"
'            RC = Cad & RC
'
'            Conn.Execute RC
'        End If
        Rs.MoveNext
    Wend
    Rs.Close
    CadenaDesdeOtroForm = ""
    
    Set Rs = Nothing
    Set miRsAux = Nothing
    
    If Opcion = 13 Then
        'Puede ser carta
        If chkCartaAbonos.Value Then
            'En nommacta pongo la provincia  (desprovi)
            cad = "INSERT INTO usuarios.zcuentas(codusu,codmacta,nommacta,razosoci,dirdatos,codposta,despobla,nifdatos)"
            cad = cad & " Select " & vUsu.Codigo & ",codmacta,desprovi,razosoci,dirdatos,codposta,despobla,nifdatos FROM cuentas WHERE "
            cad = cad & " codmacta IN (select distinct(codmacta) from usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo & ")"
            Ejecuta cad
        
        
            cad = "apoderado"
            RC = DevuelveDesdeBD("contacto", "empresa2", "1", "1", "N", cad)
            If RC = "" Then RC = cad
            If RC <> "" Then
                cad = "UPDATE usuarios.ztesoreriacomun SET observa1='" & DevNombreSQL(RC) & "'"
                cad = cad & " WHERE codusu = " & vUsu.Codigo
                Conn.Execute cad
            End If
        End If
    End If
    
    If Me.chkRem(0).Value = 1 Then
        If I = 1 Then
            MsgBox "No hay vencimientos asociados a las transferencias", vbExclamation
            Exit Function
        End If
    End If
    ListadoTransferencias = True
    Exit Function
EListadoTransferencias:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    Set miRsAux = Nothing
End Function





Private Function ListAseguBasico() As Boolean
    On Error GoTo EListAseguBasico
    ListAseguBasico = False
    
    cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = "Select * from cuentas where numpoliz<>"""""
    Sql = ""
    RC = CampoABD(Text3(21), "F", "fecsolic", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecconce", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "codmacta", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "codmacta", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    If Sql <> "" Then cad = cad & Sql
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & " ORDER BY " & RC
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,importe1,"
    cad = cad & "importe2,observa1,observa2) VALUES (" & vUsu.Codigo & ","
        
    While Not miRsAux.EOF
        CONT = CONT + 1
        Sql = CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','" & DBLet(miRsAux!nifdatos, "T") & "','"
        Sql = Sql & DevNombreSQL(miRsAux!numpoliz) & "',"
        'Fecha sol y concesion
        Sql = Sql & CampoBD_A_SQL(miRsAux!fecsolic, "F", True) & "," & CampoBD_A_SQL(miRsAux!fecconce, "F", True) & ","
        'Importes sol y concesion
        Sql = Sql & CampoBD_A_SQL(miRsAux!credisol, "N", True) & "," & CampoBD_A_SQL(miRsAux!credicon, "N", True) & ","
        'Observaciones
        RC = Memo_Leer(miRsAux!observa)
        If Len(RC) = 0 Then
            'Los dos campos NULL
            Sql = Sql & "NULL,NULL"
        Else
            If Len(RC) < 255 Then
                Sql = Sql & "'" & DevNombreSQL(RC) & "',NULL"
            Else
                Sql = Sql & "'" & DevNombreSQL(Mid(RC, 1, 255))
                RC = Mid(RC, 256)
                Sql = Sql & "','" & DevNombreSQL(Mid(RC, 1, 255)) & "'"
            End If
        End If
        
        Sql = Sql & ")"
        Conn.Execute cad & Sql
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAseguBasico = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAseguBasico:
    MuestraError Err.Number, "ListAseguBasico"
End Function





Private Function ListAsegFacturacion() As Boolean
Dim FP As Integer 'Forma de pago
Dim Cadpago As String
    On Error GoTo EListAsegFacturacion
    ListAsegFacturacion = False
    
    cad = "DELETE FROM Usuarios.zpendientes  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    
    If Me.optFecgaASig(0).Value Then
        cad = "fecfaccl"
    Else
        cad = "fecvenci"
    End If
        
    Sql = ""
    RC = CampoABD(Text3(21), "F", cad, True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(Text3(22), "F", cad, False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    
    
    
    cad = "Select scobro.*,nommacta,numpoliz,nomforpa,forpa from scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    cad = cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba
    If Sql <> "" Then cad = cad & Sql
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & " ORDER BY " & RC
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0

    cad = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,"
    cad = cad & "codforpa, nomforpa, codmacta, nombre, fecVto, importe,"
    cad = cad & "Situacion,pag_cob, vencido,  gastos) VALUES (" & vUsu.Codigo & ","
    Cadpago = ","
    While Not miRsAux.EOF
        CONT = CONT + 1
        Sql = "'" & miRsAux!NUmSerie & "','" & Format(miRsAux!codfaccl, "000000000") & "','" & Format(miRsAux!fecfaccl, FormatoFecha) & "',"
        FP = miRsAux!codforpa
        If optFP(1).Value Then
            If DBLet(miRsAux!Forpa, "N") > 0 Then
                FP = miRsAux!Forpa
                If InStr(1, Cadpago, "," & FP & ",") = 0 Then Cadpago = Cadpago & FP & ","
            End If
        End If
        Sql = Sql & miRsAux!numorden & "," & FP & ",'" & DevNombreSQL(miRsAux!nomforpa) & "','" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta)
        Sql = Sql & "','" & Format(miRsAux!FecVenci, FormatoFecha) & "',"
        'IMporte
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        Sql = Sql & TransformaComasPuntos(CStr(Importe))
        'Situacion tengo numpoliza
        Sql = Sql & ",'" & DevNombreSQL(miRsAux!numpoliz) & "',"
        'Gastos e imvenci van a la columna pag_cob   Julio 2009
        Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
        Sql = Sql & TransformaComasPuntos(CStr(Importe))
        'El resto
        Sql = Sql & ",0,NULL)"
        
        Conn.Execute cad & Sql
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If CONT = 0 Then
        MsgBox "Ningun datos con esos valores", vbExclamation
        Exit Function
    End If
    
    
    'Si ha cambiado valores en la forma de pago
    If Len(Cadpago) > 1 Then
        Cadpago = Mid(Cadpago, 2)
        Cadpago = Mid(Cadpago, 1, Len(Cadpago) - 1)
        cad = "select codforpa,nomforpa from sforpa where codforpa in (" & Cadpago & ") GROUP by  codforpa"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = " WHERE codusu = " & vUsu.Codigo & " AND codforpa = "
        While Not miRsAux.EOF
            Sql = "UPDATE Usuarios.zpendientes SET nomforpa = '" & DevNombreSQL(miRsAux!nomforpa) & "'" & cad & miRsAux!codforpa
            If Not Ejecuta(Sql) Then MsgBox "Error actualizando tmp.  Forpa: " & miRsAux!codforpa & " " & miRsAux!nomforpa, vbExclamation
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    ListAsegFacturacion = True
    
    
    Exit Function
EListAsegFacturacion:
    MuestraError Err.Number, "ListAseguBasico"
End Function



Private Function ListAsegImpagos() As Boolean
    On Error GoTo EListAsegImpagos
    ListAsegImpagos = False
    
    cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,scobro.codmacta,nommacta,despobla,desprovi,numpoliz,nomforpa from "
    cad = cad & "scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    cad = cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba
    'Impagados
    cad = cad & " AND devuelto = 1"
    Sql = ""
    RC = CampoABD(Text3(21), "F", "fecvenci", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecvenci", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    If Sql <> "" Then cad = cad & Sql
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & " ORDER BY " & RC
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,texto5,texto6,fecha1,importe1) VALUES (" & vUsu.Codigo & ","
        
    While Not miRsAux.EOF
        CONT = CONT + 1
        Sql = CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','" & DevNombreSQL(DBLet(miRsAux!desPobla, "T")) & "','"
        Sql = Sql & DevNombreSQL(DBLet(miRsAux!desProvi, "T")) & "','" & DevNombreSQL(miRsAux!numpoliz) & "','"
        Sql = Sql & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"
        'Fecha vto
        Sql = Sql & CampoBD_A_SQL(miRsAux!FecVenci, "F", True) & ","
        'Importes sol y concesion
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        Sql = Sql & TransformaComasPuntos(CStr(Importe))
        
    
        Sql = Sql & ")"
        Conn.Execute cad & Sql
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAsegImpagos = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegImpagos:
    MuestraError Err.Number, "ListAsegImpagos"
End Function


Private Function ListAsegEfectos() As Boolean
Dim TotalCred As Currency

    On Error GoTo EListAsegEfectos
    ListAsegEfectos = False
    
    cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,fecfaccl,devuelto,scobro.codmacta,nommacta,credicon from "
    cad = cad & "scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    cad = cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba

    Sql = ""
    RC = CampoABD(Text3(21), "F", "fecvenci", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecvenci", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    If Sql <> "" Then cad = cad & Sql
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & " ORDER BY codmacta,fecvenci"
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,fvto,impvto,disponible,vencida
    cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,importe1,importe2,opcion) VALUES (" & vUsu.Codigo & ","
    RC = ""
    
    While Not miRsAux.EOF
        If RC <> miRsAux!codmacta Then
            RC = miRsAux!codmacta
            TotalCred = DBLet(miRsAux!credicon, "N")
            CadenaDesdeOtroForm = ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            If IsNull(miRsAux!credicon) Then
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "0,00','"
            Else
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(miRsAux!credicon, FormatoImporte) & "','"
            End If
        End If
        CONT = CONT + 1
        Sql = CONT & CadenaDesdeOtroForm
        Sql = Sql & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"
        'Fecha fac
        Sql = Sql & CampoBD_A_SQL(miRsAux!fecfaccl, "F", True) & ","
        'Fecha vto
        Sql = Sql & CampoBD_A_SQL(miRsAux!FecVenci, "F", True) & ","
        'Importes sol y concesion
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        Sql = Sql & TransformaComasPuntos(CStr(Importe))
        TotalCred = TotalCred - Importe
        Sql = Sql & "," & TransformaComasPuntos(CStr(TotalCred))
       
        'Devuelto
        Sql = Sql & "," & DBLet(miRsAux!Devuelto, "N") & ")"
    
        Conn.Execute cad & Sql
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAsegEfectos = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegEfectos:
    MuestraError Err.Number, "ListAsegEfec"
End Function



Private Sub GeneraComboCuentas()
'
'    If Opcion = 1 Then
'        'COBROS PENDIENTES
'    Else: Pagos
'
        cmbCuentas(Opcion - 1).Clear
        cmbCuentas(Opcion - 1).AddItem "Sin especificar"
        
        cmbCuentas(Opcion - 1).AddItem "Crear selección"
              
        'En el tag tendremos las cuentas seleccionadas
        If Me.cmbCuentas(Opcion - 1).Tag <> "" Then cmbCuentas(Opcion - 1).AddItem "Cuentas seleccionadas"


    'Cargo aqui tb los checks
    CargaTextosTipoPagos False
End Sub



Private Sub CargaTextosTipoPagos(Reclamaciones As Boolean)
    
    Set miRsAux = New ADODB.Recordset
    cad = "Select tipoformapago, descformapago,siglas from stipoformapago order by tipoformapago "
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Reclamaciones Then
            chkTipPagoRec(miRsAux!tipoformapago).Caption = miRsAux!siglas
            chkTipPagoRec(miRsAux!tipoformapago).Visible = True
            chkTipPagoRec(miRsAux!tipoformapago).Value = 1
        
        Else
            chkTipPago(miRsAux!tipoformapago).Caption = miRsAux!siglas
            chkTipPago(miRsAux!tipoformapago).Visible = True
            chkTipPago(miRsAux!tipoformapago).Value = 1
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



'Para conceptos y diarios
'Opcion: 0- Diario
'        1- Conceptos
'        2- Centros de coste
'        3- Gastos fijos
'        4. Hco compensaciones
Private Sub LanzaBuscaGrid(Indice As Integer, OpcionGrid As Byte)


'--monica
'
'    Select Case OpcionGrid
'    Case 0
'    'Diario
'        DevfrmCCtas = "0"
'        cad = "Número|numdiari|N|30·"
'        cad = cad & "Descripción|desdiari|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "Tiposdiario"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "Diario"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'           Me.txtDiario(Indice) = RecuperaValor(DevfrmCCtas, 1)
'           Me.txtDescDiario(Indice) = RecuperaValor(DevfrmCCtas, 2)
'        End If
' Case 1
'        'Conceptos
'        DevfrmCCtas = "0"
'        cad = "Codigo|codconce|N|30·"
'        cad = cad & "Descripción|nomconce|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "Conceptos"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "CONCEPTOS"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'           Me.txtConcpto(Indice) = RecuperaValor(DevfrmCCtas, 1)
'           Me.txtDescConcepto(Indice) = RecuperaValor(DevfrmCCtas, 2)
'        End If
'
'    Case 2
'        'Centros de coste
'        DevfrmCCtas = "0"
'        cad = "Codigo|codccost|T|30·"
'        cad = cad & "Descripción|nomccost|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "cabccost"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "Centros de coste"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'
'           txtCCost(Indice) = RecuperaValor(DevfrmCCtas, 1)
'           txtDescCCoste(Indice) = RecuperaValor(DevfrmCCtas, 2)
'        End If
'
'    Case 3
'        'Gasto fijos  sgastfij codigo Descripcion
'        DevfrmCCtas = "0"
'        cad = "Código|codigo|T|30·"
'        cad = cad & "Descripción|Descripcion|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "sgastfij"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "Gastos fijos"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'
'           txtGastoFijo(Indice) = RecuperaValor(DevfrmCCtas, 1)
'           txtDescGastoFijo(Indice) = RecuperaValor(DevfrmCCtas, 2)
'        End If
'
'    Case 4
'        'Gasto fijos  sgastfij codigo Descripcion
'        '-------------------------------------------
'        DevfrmCCtas = "0"
'        cad = "Código|codigo|T|10·"
'        cad = cad & "Fecha|fecha|T|26·"
'        cad = cad & "Cuenta|codmacta|T|14·"
'        cad = cad & "Nombre|nommacta|T|50·"
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
'    End Select
End Sub

                                       '                Para saber el index del listview
Public Sub InsertaItemComboCompensaVto(TEXTO As String, Indice As Integer)
    Me.cboCompensaVto.AddItem TEXTO
    Me.cboCompensaVto.ItemData(Me.cboCompensaVto.NewIndex) = Indice
End Sub


Private Function GeneraDatosTalPag() As Boolean
Dim B As Boolean

    'Borramos los tmp
    Sql = "DELETE FROM usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Conn.Execute Sql

    If chkLstTalPag(3).Value = 1 Then
        B = GeneraDatosTalPagDesglosado
    Else
        'Sin desglosar
        B = GeneraDatosTalPagSinDesglose
    End If
    GeneraDatosTalPag = B
End Function

Private Function GeneraDatosTalPagDesglosado() As Boolean
    On Error GoTo EGeneraDatosTalPag
    GeneraDatosTalPagDesglosado = False
    
    

       
       
    Sql = "select slirecepdoc.*,scarecepdoc.*,nommacta,nifdatos from slirecepdoc,scarecepdoc,cuentas "
    Sql = Sql & " where slirecepdoc.id =scarecepdoc.codigo and scarecepdoc.codmacta=cuentas.codmacta"
    If Text3(24).Text <> "" Then Sql = Sql & " AND fecharec >= '" & Format(Text3(24).Text, FormatoFecha) & "'"
    If Text3(25).Text <> "" Then Sql = Sql & " AND fecharec <= '" & Format(Text3(25).Text, FormatoFecha) & "'"
    'Contabilizado
    'SQL = SQL & " AND Contabilizada =  1"
    'Si esta llevada a banco o no
    If cboListPagare.ListIndex >= 1 Then Sql = Sql & " AND LlevadoBanco = " & Abs(cboListPagare.ListIndex = 1)
    '------------------------------------------------------------------------
    I = -1
    If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
        'Solo uno seleccionado
        I = 0
        If (chkLstTalPag(0).Value = 1) Then I = 1
            
    Else
        If (chkLstTalPag(0).Value = 0) Then
            MsgBox "Seleccione talon o pagare", vbExclamation
            Exit Function
        End If
    End If
    If I >= 0 Then Sql = Sql & " AND talon = " & I

    'Si ID
    If txtNumFac(2).Text <> "" Then Sql = Sql & " AND codigo >= " & txtNumFac(2).Text
    If txtNumFac(3).Text <> "" Then Sql = Sql & " AND codigo <= " & txtNumFac(3).Text

    Set Rs = New ADODB.Recordset
    
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    RC = ""
    While Not Rs.EOF
        I = I + 1
        'ztesoreriacomun (`codusu`,`codigo
        Sql = ", (" & vUsu.Codigo & "," & I & ","
        
        'texto1`,`texto2`,`texto3`,y el 4
        Sql = Sql & "'" & DevNombreSQL(Rs!numeroref) & "','" & DevNombreSQL(Rs!banco) & "','"
        Sql = Sql & DevNombreSQL(Rs!codmacta) & "','" & DevNombreSQL(Rs!Nommacta) & "','"
        
        
        '5 Serie y numero factura
        Sql = Sql & Rs!NUmSerie & Format(Rs!numfaccl, "000000") & "',"
        '`importe1`
        'SQL = SQL & TransformaComasPuntos(CStr(Rs!Importe)) & ","
        Sql = Sql & TransformaComasPuntos(CStr(Rs.Fields(5))) & ",'"   'La columna 5 es sli.importe
        
        'texto6=nifdatos
        Sql = Sql & DevNombreSQL(DBLet(Rs!nifdatos, "N"))
        
        '`fecha1`,`fecha2`,`fecha3`
        Sql = Sql & "','" & Format(Rs!fecharec, FormatoFecha) & "',"
        Sql = Sql & "'" & Format(Rs!fechavto, FormatoFecha) & "',"
        Sql = Sql & "'" & Format(Rs!fecfaccl, FormatoFecha) & "')"
    
        RC = RC & Sql
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    If I > 0 Then
        RC = Mid(RC, 3) 'QUITO LA PRIMERA COMA
        'OK hay datos. Insertamos
        Sql = "INSERT INTO usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`texto3`,"
        Sql = Sql & "`texto4`,`texto5`,`importe1`,texto6,`fecha1`,`fecha2`,`fecha3`) VALUES "
        Sql = Sql & RC
        Conn.Execute Sql
    
        'Si estamos emitiendo el justicante de recepcion, guardare en z340 los campos
        'fiscales del cliente para su impresion
        If Me.chkLstTalPag(2).Value = 1 Then
            Sql = "DELETE FROM usuarios.z347 WHERE codusu = " & vUsu.Codigo
            Conn.Execute Sql
            
            Sql = "DELETE FROm Usuarios.z347carta WHERE codusu = " & vUsu.Codigo
            Conn.Execute Sql
            
            espera 0.3
            
            
            'En texto3 esta la codmacta
            Sql = "Select texto3 from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " GROUP BY texto3"
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            RC = ""
            While Not Rs.EOF
                RC = RC & ", '" & Rs!texto3 & "'"
                Rs.MoveNext
            Wend
            Rs.Close
            
            
            
            
            
            'No puede ser EOF
            RC = Trim(Mid(RC, 2))
            'Monto un superselect
            'pongo el IGNORE por si acaso hay cuentas con el mismo NIF
            Sql = "insert ignore into usuarios.z347 (`codusu`,`cliprov`,`nif`,`razosoci`,`dirdatos`,`codposta`,`despobla`,`Provincia`)"
            Sql = Sql & " SELECT " & vUsu.Codigo & ",0,nifdatos,razosoci,dirdatos,codposta,despobla,desprovi FROM cuentas where codmacta in (" & RC & ")"
            Conn.Execute Sql
    
    
    
            'Ahora meto los datos de la empresa
            cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir,"
            cad = cad & "contacto) VALUES ("
            cad = cad & vUsu.Codigo
                
                
            'Monta Datos Empresa
            Rs.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
            If Rs.EOF Then
                MsgBox "Error en los datos de la empresa " & vEmpresa.nomempre
                RC = ",'','','','','',''"  '6 campos
            Else
                RC = DBLet(Rs!siglasvia) & " " & DBLet(Rs!Direccion) & "  " & DBLet(Rs!numero) & ", " & DBLet(Rs!puerta)
                RC = ",'" & DBLet(Rs!nifempre) & "','" & vEmpresa.nomempre & "','" & RC & "','"
                RC = RC & DBLet(Rs!codpos) & "','" & DBLet(Rs!Poblacion) & "','" & DBLet(Rs!provincia) & "','" & DBLet(Rs!contacto) & "')"
            End If
            Rs.Close
            cad = cad & RC
            Conn.Execute cad
            
            
            
    
        End If
        GeneraDatosTalPagDesglosado = True
    Else
        'I>0
        MsgBox "No hay datos", vbExclamation
    End If

EGeneraDatosTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function



Private Function GeneraDatosTalPagSinDesglose() As Boolean
    On Error GoTo EGeneraDatosTalPag
    GeneraDatosTalPagSinDesglose = False
    
    

       
       
    Sql = "select scarecepdoc.*,nommacta from scarecepdoc,cuentas "
    Sql = Sql & " where  scarecepdoc.codmacta=cuentas.codmacta"
    If Text3(24).Text <> "" Then Sql = Sql & " AND fecharec >= '" & Format(Text3(24).Text, FormatoFecha) & "'"
    If Text3(25).Text <> "" Then Sql = Sql & " AND fecharec <= '" & Format(Text3(25).Text, FormatoFecha) & "'"
    'Contabilizado
    'SQL = SQL & " AND Contabilizada =  1"
    'Si esta llevada a banco o no
    'SQL = SQL & " AND LlevadoBanco = " & Abs(chkLstTalPag(2).Value)
    If cboListPagare.ListIndex >= 1 Then Sql = Sql & " AND LlevadoBanco = " & Abs(cboListPagare.ListIndex = 1)
    
    I = -1
    If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
        'Solo uno seleccionado
        I = 0
        If (chkLstTalPag(0).Value = 1) Then I = 1
            
    Else
        If (chkLstTalPag(0).Value = 0) Then
            MsgBox "Seleccione talon o pagare", vbExclamation
            Exit Function
        End If
    End If
    If I >= 0 Then Sql = Sql & " AND talon = " & I
    'Si ID
    If txtNumFac(2).Text <> "" Then Sql = Sql & " AND codigo >= " & txtNumFac(2).Text
    If txtNumFac(3).Text <> "" Then Sql = Sql & " AND codigo <= " & txtNumFac(3).Text



    Set Rs = New ADODB.Recordset
    
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    RC = ""
    While Not Rs.EOF
        I = I + 1
        'ztesoreriacomun (`codusu`,`codigo
        Sql = ", (" & vUsu.Codigo & "," & I & ","
        
        'texto1`,`texto2`,`texto3`,y el 4
        Sql = Sql & "'" & DevNombreSQL(Rs!numeroref) & "','" & DevNombreSQL(Rs!banco) & "','"
        Sql = Sql & DevNombreSQL(Rs!codmacta) & "','" & DevNombreSQL(Rs!Nommacta) & "','"
        
        
        '5 Serie y numero factura
        Sql = Sql & "',"
        '`importe1`
        'SQL = SQL & TransformaComasPuntos(CStr(Rs.Fields(5))) & ","   'La columna 5 es sli.importe
        Sql = Sql & TransformaComasPuntos(CStr(Rs!Importe)) & ","
        
        '
        '`fecha1`,`fecha2`,`fecha3`
        Sql = Sql & "'" & Format(Rs!fecharec, FormatoFecha) & "',"
        Sql = Sql & "'" & Format(Rs!fechavto, FormatoFecha) & "',"
        Sql = Sql & "'" & Format(Now, FormatoFecha) & "')"
    
        RC = RC & Sql
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    If I > 0 Then
        RC = Mid(RC, 3) 'QUITO LA PRIMERA COMA
        'OK hay datos. Insertamos
        Sql = "INSERT INTO usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`texto3`,"
        Sql = Sql & "`texto4`,`texto5`,`importe1`,`fecha1`,`fecha2`,`fecha3`) VALUES "
        Sql = Sql & RC
        Conn.Execute Sql
        GeneraDatosTalPagSinDesglose = True
    Else
        MsgBox "No hay datos", vbExclamation
    End If
    
    

EGeneraDatosTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function





Private Function ListadoOrdenPago() As Boolean
    On Error GoTo EListadoOrdenPago
    ListadoOrdenPago = False

    'Borramos
    cad = "DELETE from usuarios.zlistadopagos WHERE codusu = " & vUsu.Codigo
    Conn.Execute cad
    Set miRsAux = New ADODB.Recordset
    'Inserttamos
    RC = ""
    If txtCtaBanc(3).Text <> "" Or txtCtaBanc(4).Text <> "" Then
        If txtCtaBanc(3).Text <> "" Then RC = " codmacta >= '" & txtCtaBanc(3).Text & "'"
        
        If txtCtaBanc(4).Text <> "" Then
            If RC <> "" Then RC = RC & " AND "
            RC = RC & " codmacta <= '" & txtCtaBanc(4).Text & "'"
        End If
        cad = "Select codmacta from ctabancaria where " & RC
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        RC = ""
        While Not miRsAux.EOF
            RC = RC & ", '" & miRsAux!codmacta & "'"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If RC = "" Then
            MsgBox "Ningún banco con esos valores", vbExclamation
            Exit Function
        End If
           
        RC = Mid(RC, 2)
    End If
    
    
    Sql = ""
    If Text3(26).Text <> "" Then Sql = Sql & " AND fecefect >= '" & Format(Text3(26).Text, FormatoFecha) & "'"
    If Text3(27).Text <> "" Then Sql = Sql & " AND fecefect <= '" & Format(Text3(27).Text, FormatoFecha) & "'"
    If RC <> "" Then Sql = Sql & " AND ctabanc1 in (" & RC & ")"
    
    
    'JULIO2013
    'La variable contdocu grabaremos emitdocum, y en el listado sabremos si es de talon/pagere
    'para poder separalos
    'Antes: contdocu, ahora emitdocum
    
    'Agosto 2014
    'Tipo de pago
    cad = "select " & vUsu.Codigo & ",`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,`impefect`-coalesce(imppagad,0),"
    cad = cad & " `ctabanc1`,`ctabanc2`,`emitdocum`,spagop.entidad,spagop.oficina,spagop.CC,spagop.cuentaba,"
    cad = cad & " nommacta,'error','error',descformapago "
    
    cad = cad & " from spagop,cuentas,sforpa,stipoformapago "
    cad = cad & " WHERE spagop.ctaprove = cuentas.codmacta AND spagop.codforpa=sforpa.codforpa and tipoformapago=tipforpa"
    'Ponemos un check si salen negativos o no
    RC = " AND impefect >=0"
    If Me.chkPagBanco(0).Value = 1 And Me.chkPagBanco(1).Value = 1 Then RC = "" 'Salen todos
    cad = cad & RC 'todos o solo positivos
    cad = cad & Sql
    
    Sql = "INSERT INTO usuarios.zlistadopagos (`codusu`,`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,"
    Sql = Sql & " `impefect`,`ctabanc1`,`ctabanc2`,`contdocu`,`entidad`,`oficina`,`CC`,`cuentaba`,"
    Sql = Sql & " `nomprove`,`nombanco`,`cuentabanco`,TipoForpa) " & cad
    Conn.Execute Sql
    
    cad = DevuelveDesdeBD("count(*)", "usuarios.zlistadopagos", "codusu", vUsu.Codigo)
    If Val(cad) = 0 Then
        MsgBox "Ningun vencimiento con esos valores", vbExclamation
        Exit Function
    End If
    
    'Actualizo los datos de los bancos `nombanco`,`cuentabanco`
    cad = "Select ctabanc1 from usuarios.zlistadopagos WHERE codusu = " & vUsu.Codigo
    cad = cad & " AND ctabanc1 <>'' GROUP BY ctabanc1"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF
        cad = cad & miRsAux!ctabanc1 & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    While cad <> ""
        I = InStr(1, cad, "|")
        If I = 0 Then
            cad = ""
        Else
            RC = Mid(cad, 1, I - 1)
            cad = Mid(cad, I + 1)
            
            Sql = "Select ctabancaria.codmacta,ctabancaria.entidad, ctabancaria.oficina, ctabancaria.control, ctabancaria.ctabanco,"
            Sql = Sql & " ctabancaria.descripcion,nommacta from  ctabancaria,cuentas where ctabancaria.codmacta=cuentas.codmacta "
            Sql = Sql & " AND ctabancaria.codmacta ='" & RC & "'"
            miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                Sql = "Cuenta banco erronea: " & vbCrLf & "Hay vencimientos asociados a la cuenta " & RC & " que no esta en bancos"
                MsgBox Sql, vbExclamation
            Else
                Sql = DBLet(miRsAux!Descripcion, "T")
                If Sql = "" Then Sql = miRsAux!Nommacta
                Sql = DevNombreSQL(Sql) & "|"
                
                'Enti8dad...
                I = DBLet(miRsAux!Entidad, "0")
                Sql = Sql & Format(I, "0000")
                                'Oficina...
                I = DBLet(miRsAux!Oficina, "0")
                Sql = Sql & Format(I, "0000")
                                'CC...
                RC = DBLet(miRsAux!Control, "T")
                If RC = "" Then RC = "**"
                Sql = Sql & RC
                'cuenta
                RC = DBLet(miRsAux!CtaBanco, "T")
                If RC = "" Then RC = "    **"
                Sql = Sql & RC & "|"
                
                
                RC = "UPDATE usuarios.zlistadopagos set `nombanco`='" & RecuperaValor(Sql, 1)
                RC = RC & "',`cuentabanco`='" & RecuperaValor(Sql, 2) & "'"
                RC = RC & " WHERE ctabanc1 = '" & miRsAux!codmacta & "' AND codusu = " & vUsu.Codigo
                Conn.Execute RC
                
            End If
            miRsAux.Close
        End If
    Wend
    
    ListadoOrdenPago = True
    Set miRsAux = Nothing
    Exit Function
EListadoOrdenPago:
    MuestraError Err.Number, "ListadoOrdenPago"
End Function



Private Function ListadoReclamas() As Boolean

On Error GoTo EListadoReclamas

    ListadoReclamas = False
        

    Sql = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = ""
    cad = ""
    
    If Text3(28).Text <> "" Or Text3(29).Text <> "" Then
        RC = DesdeHasta("F", 28, 29, "F.Reclama")
        If RC <> "" Then cad = cad & " " & RC
            
        RC = CampoABD(Text3(28), "F", "fecreclama", True)
        If RC <> "" Then
            If Sql <> "" Then Sql = Sql & " AND "
            Sql = Sql & RC
        End If
        
        RC = CampoABD(Text3(29), "F", "fecreclama", False)
        If RC <> "" Then
            If Sql <> "" Then Sql = Sql & " AND "
            Sql = Sql & RC
        End If
        
    End If
    
    
    If txtCta(15).Text <> "" Or txtCta(16).Text <> "" Then
        RC = DesdeHasta("C", 15, 16, "Cta")
        If RC <> "" Then cad = cad & " " & RC
            
        RC = CampoABD(txtCta(15), "T", "codmacta", True)
        If RC <> "" Then
            If Sql <> "" Then Sql = Sql & " AND "
            Sql = Sql & RC
        End If
        
        RC = CampoABD(txtCta(16), "T", "codmacta", False)
        If RC <> "" Then
            If Sql <> "" Then Sql = Sql & " AND "
            Sql = Sql & RC
        End If
        
    End If
    If Sql <> "" Then Sql = " WHERE " & Sql
    Sql = "Select * from shcocob" & Sql
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    RC = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`"
    RC = RC & "texto3`,`texto4`,`texto5`,`texto6`,`importe1`,`importe2`,`fecha1`,`fecha2`,"
    RC = RC & "`fecha3`,`texto`,`observa2`,`opcion`) VALUES "
    Sql = ""
    While Not Rs.EOF
        NumRegElim = NumRegElim + 1
        Sql = Sql & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & Rs!codmacta & "','"
        'text 2 y 3
        Sql = Sql & DevNombreSQL(Rs!Nommacta) & "','" & Rs!NUmSerie & Format(Rs!codfaccl, "000000") & "','"
        '4 y 5
        Sql = Sql & Rs!numorden & "','"
        If Val(Rs!carta) = 1 Then
            Sql = Sql & "Email"
        ElseIf Val(Rs!carta) = 2 Then
            Sql = Sql & "Teléfono"
        Else
            Sql = Sql & "Carta"
        End If
        'Text6, importe 1 y 2
        Sql = Sql & "',''," & TransformaComasPuntos(CStr(Rs!ImpVenci)) & ",NULL,"
        'Fec1 reclama fec2 factra   fec3
        Sql = Sql & "'" & Format(Rs!Fecreclama, FormatoFecha) & "','" & Format(Rs!fecfaccl, FormatoFecha) & "',NULL,"
        DevfrmCCtas = Memo_Leer(Rs!observaciones)
        If DevfrmCCtas = "" Then
            DevfrmCCtas = "NULL"
        Else
            DevfrmCCtas = "'" & DevNombreSQL(DevfrmCCtas) & "'"
        End If
        Sql = Sql & DevfrmCCtas & ",NULL,0)"


        'Siguiente
        Rs.MoveNext
        
        
        If Len(Sql) > 100000 Then
            Sql = Mid(Sql, 2) 'QUITO LA COMA
            Sql = RC & Sql
            Conn.Execute Sql
            Sql = ""
        End If
            
        
    Wend
    Rs.Close
        If Sql <> "" Then
            Sql = Mid(Sql, 2) 'QUITO LA COMA
            Sql = RC & Sql
            Conn.Execute Sql
        End If
        
        
    If NumRegElim > 0 Then
        ListadoReclamas = True
    Else
        MsgBox "Ningun dato devuelto", vbExclamation
    End If
    Exit Function
EListadoReclamas:
    MuestraError Err.Number, Err.Description
End Function





'******************************************************************************************
'
'   Listado gastos fijos

Private Function ListadoGastosFijos() As Boolean

On Error GoTo EListadoGF

    ListadoGastosFijos = False
        

    Sql = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = ""
    cad = ""
    
    
    DevfrmCCtas = "" ' ON del left join , NO al WHERE
    If Text3(30).Text <> "" Or Text3(31).Text <> "" Then
        RC = DesdeHasta("F", 30, 31, "Fecha")
        If RC <> "" Then cad = cad & " " & Trim(RC)
            
        RC = CampoABD(Text3(30), "F", "fecha", True)
        If RC <> "" Then
            If Sql <> "" Then Sql = Sql & " AND "
            Sql = Sql & RC
        End If
        
        RC = CampoABD(Text3(31), "F", "fecha", False)
        If RC <> "" Then
            If Sql <> "" Then Sql = Sql & " AND "
            Sql = Sql & RC
        End If
        
    End If
    DevfrmCCtas = Sql
    Sql = ""
    
    'Este si que va al where
    If txtGastoFijo(0).Text <> "" Or txtGastoFijo(1).Text <> "" Then
        RC = DesdeHasta("GF", 0, 1, "Codigo")
        If RC <> "" Then
            If cad <> "" Then
                'Ya esta la fecha
                If Len(cad & RC) > 100 Then cad = cad & """ + chr(13) + """
            End If
            cad = cad & " " & Trim(RC)
        End If
            
        RC = CampoABD(txtGastoFijo(0), "N", "sgastfij.codigo", True)
        If RC <> "" Then
            If Sql <> "" Then Sql = Sql & " AND "
            Sql = Sql & RC
        End If
        
        RC = CampoABD(txtGastoFijo(1), "N", "sgastfij.codigo", False)
        If RC <> "" Then
            If Sql <> "" Then Sql = Sql & " AND "
            Sql = Sql & RC
        End If
        
    End If
    
   
   
    RC = " FROM sgastfij left join sgastfijd ON sgastfij.Codigo = sgastfijd.Codigo"
    If DevfrmCCtas <> "" Then RC = RC & " AND " & DevfrmCCtas
    If Sql <> "" Then RC = RC & " WHERE " & Sql
    Sql = "SELECT sgastfij.codigo,descripcion,ctaprevista,fecha,importe" & RC
    
    

    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    RC = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`"
    RC = RC & "texto3`,`importe1`,`fecha1`) VALUES "
    Sql = ""
    While Not Rs.EOF
        NumRegElim = NumRegElim + 1
        Sql = Sql & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & Format(Rs!Codigo, "00000") & "','"
        'text 2 y 3
        Sql = Sql & DevNombreSQL(Rs!Descripcion) & "','" & Rs!Ctaprevista & "',"
       
  
        'Detalla
        If IsNull(Rs!Fecha) Then
            Sql = Sql & "0,'" & Format(Now, FormatoFecha) & "'"
        Else
            Sql = Sql & TransformaComasPuntos(DBLet(Rs!Importe, "N")) & ",'" & Format(Rs!Fecha, FormatoFecha) & "'"
        End If
        Sql = Sql & ")"
        
        'Siguiente
        Rs.MoveNext
            
        
    Wend
    Rs.Close
    If Sql <> "" Then
        Sql = Mid(Sql, 2) 'QUITO LA COMA
        Sql = RC & Sql
        Conn.Execute Sql
    End If
        
        
    If NumRegElim = 0 Then
        MsgBox "Ningun dato devuelto", vbExclamation
        Exit Function
    End If
    
    
    'Updateo la cuenta bancaria
    RC = "Select texto3 from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " GROUP BY 1"
    Rs.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not Rs.EOF
        Sql = Sql & Rs!texto3 & "|"
        Rs.MoveNext
    Wend
    Rs.Close
    
    While Sql <> ""
        NumRegElim = InStr(1, Sql, "|")
        If NumRegElim = 0 Then
            Sql = ""
        Else
            RC = Mid(Sql, 1, NumRegElim - 1)
            Sql = Mid(Sql, NumRegElim + 1)
            
            RC = "Select codmacta,nommacta from cuentas where codmacta='" & RC & "'"
            Rs.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                RC = "UPDATE usuarios.ztesoreriacomun SET texto4='" & DevNombreSQL(Rs!Nommacta) & "' WHERE codusu =" & vUsu.Codigo & " AND texto3='" & Rs!codmacta & "'"
                Conn.Execute RC
            End If
            Rs.Close
        End If
    Wend
    ListadoGastosFijos = True
    Exit Function
EListadoGF:
    MuestraError Err.Number, Err.Description
End Function






'Listadoas aseguradoas
Private Function AvisosAseguradora() As Boolean



    On Error GoTo EListAsegEfectos
    AvisosAseguradora = False
    
    cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    'feccomunica,fecprorroga,fecsiniestro
    Sql = ""
    If Me.optAsegAvisos(0).Value Then
        cad = "feccomunica"
    ElseIf Me.optAsegAvisos(1).Value Then
        cad = "fecprorroga"
    Else
        cad = "fecsiniestro"
    End If
    RC = CampoABD(Text3(21), "F", cad, True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(Text3(22), "F", cad, False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then Sql = Sql & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then Sql = Sql & " AND " & RC
    
    'Significa que no ha puesto fechas
    If InStr(1, Sql, cad) = 0 Then Sql = Sql & " AND " & cad & ">='1900-01-01'"
    
    'ORDENACION
    If Me.optAsegAvisos(0).Value Then
        RC = "feccomunica"
    ElseIf Me.optAsegAvisos(1).Value Then
        RC = "fecprorroga"
    Else
        RC = "fecsiniestro"
    End If
    
    cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,fecfaccl,devuelto,scobro.codmacta,nommacta,numpoliz"
    cad = cad & ",credicon," & RC & " LaFecha" 'alias
    cad = cad & "  FROM scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    cad = cad & " and scobro.codforpa=sforpa.codforpa "
    If Sql <> "" Then cad = cad & Sql
    
    
    
    

    cad = cad & " ORDER BY " & RC & ","
    'ORDENACION 2º nivel
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & RC
    
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,faviso,fvto,impvto,disponible,vencida
    cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,fecha3,importe1,importe2,opcion) VALUES "
    RC = ""
    
    While Not miRsAux.EOF
        If Len(RC) > 500 Then
            RC = Mid(RC, 2)
            Conn.Execute cad & RC
            RC = ""
        End If
        CONT = CONT + 1
        Sql = ", (" & vUsu.Codigo & "," & CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
        Sql = Sql & DevNombreSQL(miRsAux!numpoliz) & "'"
        Sql = Sql & ",'" & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"  'texto4
        'Fecha fac
        Sql = Sql & CampoBD_A_SQL(miRsAux!fecfaccl, "F", True) & ","
        'Fecha aviso
        Sql = Sql & CampoBD_A_SQL(miRsAux!lafecha, "F", True) & ","
        'Fecha vto
        Sql = Sql & CampoBD_A_SQL(miRsAux!FecVenci, "F", True)
        
        Sql = Sql & "," & TransformaComasPuntos(CStr(miRsAux!ImpVenci))
        Sql = Sql & "," & TransformaComasPuntos(CStr(DBLet(miRsAux!Gastos, "N")))
        'Devuelto
        Sql = Sql & "," & DBLet(miRsAux!Devuelto, "N") & ")"
    
        RC = RC & Sql
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If RC <> "" Then
        RC = Mid(RC, 2)
        Conn.Execute cad & RC
    End If
    
    
    If CONT > 0 Then
        AvisosAseguradora = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegEfectos:
    MuestraError Err.Number, "Avisos aseguradoras"
End Function



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
    cad = "Select scobro.*,nomforpa from scobro,sforpa where scobro.codforpa=sforpa.codforpa "
    cad = cad & " AND codmacta = '" & Me.txtCta(17).Text & "'"
    cad = cad & " AND (transfer =0 or transfer is null) and codrem is null"
    cad = cad & " and estacaja=0 and recedocu=0"
    cad = cad & " ORDER BY fecvenci"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCompenCli.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!codfaccl, "000000")
        IT.SubItems(2) = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
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

        cad = DevuelveDesdeBD("max(codigo)", "scompenclicab", "1", "1")
        If cad = "" Then cad = "0"
        CONT = Val(cad) + 1 'ID de la operacion
        
        cad = "INSERT INTO scompenclicab(codigo,fecha,login,PC,codmacta,nommacta) VALUES (" & CONT
        cad = cad & ",now(),'" & DevNombreSQL(vUsu.Login) & "','" & DevNombreSQL(vUsu.PC)
        cad = cad & "','" & txtCta(17).Text & "','" & DevNombreSQL(DtxtCta(17).Text) & "')"
        
        Set miRsAux = New ADODB.Recordset
        Borras = True
        If Ejecuta(cad) Then
            
            Borras = Not RealizarProcesoCompensacionAbonos
        
        End If


        Set miRsAux = Nothing
        If Borras Then
            Conn.Execute "DELETE FROM scompenclilin WHERE codigo = " & CONT
            Conn.Execute "DELETE FROM scompenclicab WHERE codigo = " & CONT
            
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
        Sql = "INSERT INTO usuarios.ztmpfaclin(`codusu`,`codigo`,`Numfac`,`Fecha`,`cta`,`Cliente`,`NIF`,`Imponible`,`ImpIVA`,`retencion`,`Total`,`IVA`,TipoIva)"
        Sql = Sql & "select " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,"
        Sql = Sql & "concat(numserie,right(concat(""000000"",codfaccl),8)) fecha,date_format(fecfaccl,'%d/%m/%Y') ffaccl,"
        Sql = Sql & "scompenclilin.codmacta,if (nommacta is null,nomclien,nommacta) nomcli,"
        Sql = Sql & "date_format(fecvenci,'%d/%m/%Y') venci,impvenci,gastos,impcobro,"
        Sql = Sql & "impvenci + coalesce(gastos,0) + coalesce(impcobro,0)  tot_al"
        Sql = Sql & ",if(fecultco is null,null,date_format(fecultco,'%d/&m')) fecco ,destino"
        Sql = Sql & " From (scompenclilin left join cuentas on scompenclilin.codmacta=cuentas.codmacta)"
        Sql = Sql & ",(SELECT @rownum:=0) r WHERE codigo=" & CONT & " order by destino desc,numserie,codfaccl"
        Conn.Execute Sql
            
        
            
        
   
    
    
        
    
    
    
    
        'Datos carta
        'Datos basicos de la empresa para la carta
        cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, "
        cad = cad & "parrafo1, parrafo2, contacto, despedida,saludos,parrafo3, parrafo4, parrafo5, Asunto, Referencia)"
        cad = cad & " VALUES (" & vUsu.Codigo & ", "
        
        'Estos datos ya veremos com, y cuadno los relleno
        Set miRsAux = New ADODB.Recordset
        Sql = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia,contacto from empresa2"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'Paarafo1 Parrafo2 contacto
        Sql = "'','',''"
        'sql= "'1234567890A','Ariadna Software ','Franco Tormo 3, Bajo Izda','46007','Valencia'"
        Sql = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'," & Sql
        If Not miRsAux.EOF Then
            Sql = ""
            For I = 1 To 6
                Sql = Sql & DBLet(miRsAux.Fields(I), "T") & " "
            Next I
            Sql = Trim(Sql)
            Sql = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(Sql) & "'"
            Sql = Sql & ",'" & DBLet(miRsAux!codpos, "T") & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"

            'Contaccto
            Sql = Sql & ",NULL,NULL,'" & DevNombreSQL(DBLet(miRsAux!contacto)) & "' "
        End If
        miRsAux.Close
      
        cad = cad & Sql
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
        Sql = DBLet(miRsAux!codposta)
        If Sql <> "" Then Sql = Sql & " - "
        Sql = Sql & DevNombreSQL(CStr(DBLet(miRsAux!desPobla)))
        cad = cad & ",'" & Sql & "'"
        'Provincia
        cad = cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!desProvi))) & "'"
        miRsAux.Close
        

        
        Sql = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4,texto5,texto6, observa1, "
        Sql = Sql & "importe1, importe2, fecha1, fecha2, fecha3, observa2, opcion)"
        Sql = Sql & " VALUES (" & vUsu.Codigo & ",1,'',''," & cad
        
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
        Sql = Sql & ",'" & cad & "'," & TransformaComasPuntos(CStr(Importe))
        
        
        'importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        For I = 1 To 6
            Sql = Sql & ",NULL"
        Next
        Sql = Sql & ")"
        Conn.Execute Sql
        
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
    
    RC = "Select " & cad & " FROM scobro where (numserie,codfaccl,fecfaccl,numorden) IN (" & Sql & ")"
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
            If miRsAux!codfaccl = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(1)) Then
                If Format(miRsAux!fecfaccl, "dd/mm/yyyy") = Me.lwCompenCli.ListItems(NumRegElim).SubItems(2) Then
                    If miRsAux!numorden = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(3)) Then Destino = 1
                End If
            End If
        End If
        
        RC = "INSERT INTO scompenclilin (codigo,linea,destino," & cad & ") VALUES (" & CONT & "," & I & "," & Destino & "," & RC & ")"
        Conn.Execute RC
        
        'Para las observaciones de despues
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Destino = 0 Then 'El destino
            DevfrmCCtas = DevfrmCCtas & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000") & " " & Format(miRsAux!fecfaccl, "dd/mm/yy")
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
        'BORRAREMOS LOS VENCIMIENTOS QUE NO SEAN DESTINO a no ser que el importe restante sea 0
        Destino = 1
        If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then Destino = 0
        SQLVtosSeleccionadosCompensacion 0, Destino = 1  'sin o con el destino
        RC = "DELETE FROM scobro where (numserie,codfaccl,fecfaccl,numorden) IN (" & Sql & ")"
        
        'Para saber si ha ido bien
        Destino = 0    '0 mal,1 bien
        If Ejecuta(RC) Then
            If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then
                Destino = 1
            Else
                'Updatearemos los campos csb del vto restante. A partir del segundo
                'La variable CadenaDesdeOtroForm  tiene los que vamos a actualizar
                
                cad = ""
                J = 0
                Sql = ""
                
                Do
                    I = InStr(1, DevfrmCCtas, "|")
                    If I = 0 Then
                        DevfrmCCtas = ""
                    Else
                        RC = Mid(DevfrmCCtas, 1, I - 1)
                        If Len(RC) > 40 Then RC = Mid(RC, 1, 40)
                        DevfrmCCtas = Mid(DevfrmCCtas, I + 1)
                        J = J + 1
                        'Antes desde aqui cogia el campo
                        'Ahora desde CadenaDesdeOtroForm que tiene los campos libres
                        'Cad = RecuperaValor("text41csb|text42csb|text43csb|text51csb|text52csb|text53csb|text61csb|text62csb|text63csb|text71csb|text72csb|text73csb|text81csb|text82csb|text83csb|", J)
                        cad = RecuperaValor(CadenaDesdeOtroForm, J)
                        Sql = Sql & ", " & cad & " = '" & DevNombreSQL(RC) & "'"
                
                    End If
                Loop Until DevfrmCCtas = ""
                Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
                RC = "gastos=null, impcobro=null,fecultco=null,impvenci=" & TransformaComasPuntos(CStr(Importe))
                Sql = RC & Sql
                Sql = "UPDATE scobro SET " & Sql
                'WHERE
                RC = ""
                For J = 1 To Me.lwCompenCli.ListItems.Count
                    If Me.lwCompenCli.ListItems(J).Bold Then
                        'Este es el destino
                        RC = "NUmSerie = '" & Me.lwCompenCli.ListItems(J).Text
                        RC = RC & "' AND codfaccl = " & Val(Me.lwCompenCli.ListItems(J).SubItems(1))
                        RC = RC & " AND fecfaccl = '" & Format(Me.lwCompenCli.ListItems(J).SubItems(2), FormatoFecha)
                        RC = RC & "' AND numorden = " & Val(Me.lwCompenCli.ListItems(J).SubItems(3))
                        Exit For
                    End If
                Next
                If RC <> "" Then
                    cad = Sql & " WHERE " & RC
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

Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    Sql = ""
    For I = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(I).Checked Then
        
            Insertar = True
            If Me.lwCompenCli.ListItems(I).Bold Then
                RegistroDestino = I
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                Sql = Sql & ", ('" & lwCompenCli.ListItems(I).Text & "'," & lwCompenCli.ListItems(I).SubItems(1)
                Sql = Sql & ",'" & Format(lwCompenCli.ListItems(I).SubItems(2), FormatoFecha) & "'," & lwCompenCli.ListItems(I).SubItems(3) & ")"
            End If
            
        End If
    Next
    Sql = Mid(Sql, 2)
            
End Sub


Private Sub FijaCadenaSQLCobrosCompen()

    cad = "NUmSerie , codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, ctabanc1,"
    cad = cad & "codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, emitdocum, "
    cad = cad & "recedocu, contdocu, text33csb, text41csb, text42csb, text43csb, text51csb, text52csb,"
    cad = cad & "text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb,"
    cad = cad & "text82csb, text83csb, ultimareclamacion, agente, departamento, tiporem, CodRem, AnyoRem,"
    cad = cad & "siturem, Gastos, Devuelto, situacionjuri, noremesar, obs, transfer, estacaja, referencia,"
    cad = cad & "reftalonpag, nomclien, domclien, pobclien, cpclien, proclien, referencia1, referencia2,"
    cad = cad & "feccomunica, fecprorroga, fecsiniestro"
    
End Sub


'******************************************************************************
'******************************************************************************
'
'******************************************************************************
'******************************************************************************



Private Function ComunicaDatosSeguro_() As Boolean
Dim k As Integer

    ComunicaDatosSeguro_ = False
    
   
    NumRegElim = 0
    
    For k = 1 To Me.ListView3.ListItems.Count
        If Me.ListView3.ListItems(k).Checked Then
            DatosSeguroUnaEmpresa CInt(ListView3.ListItems(k).Tag)
      
            Sql = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
            If Sql <> "" Then NumRegElim = Val(Sql)
        End If
    Next
    
    
    
    If NumRegElim > 0 Then
        Sql = "DELETE from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
        Sql = Sql & " AND importe1<=0"
        
        
        
    
    
        '   Conn.Execute SQL
        Sql = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
        If Sql <> "" Then
            NumRegElim = Val(Sql)
        Else
            NumRegElim = 0
        End If
        
        
        ComunicaDatosSeguro_ = NumRegElim > 0
        If NumRegElim > 0 Then
            Sql = "Select texto5 from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not miRsAux.EOF
                Sql = miRsAux!texto5
                If Sql = "" Then
                    Sql = "ESPAÑA"
                Else
                    If InStr(1, Sql, " ") > 0 Then
                        Sql = Mid(Sql, 3)
                    Else
                        Sql = "" 'no updateamos
                    End If
                End If
                If Sql <> "" Then
                    Sql = "UPDATE Usuarios.ztesoreriacomun set texto5='" & DevNombreSQL(Sql) & "' WHERE codusu ="
                    Sql = Sql & vUsu.Codigo & " AND texto5='" & DevNombreSQL(miRsAux!texto5) & "'"
                    Conn.Execute Sql
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
        
    End If
End Function

Private Sub DatosSeguroUnaEmpresa(NumConta As Integer)

    
    'select numpoliz,nifdatos,numserie,codfaccl,nommacta,impvenci,gastos,impcobro,credicon from scobro,cuentas where
    'scobro.codmacta = cuentas.codmacta     fecbajcre
    
    'JUlio2013
    'Para fontenas iran por PAIS
    'añadiremos en text05 el pais
    Sql = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4,"
    Sql = Sql & " importe1,  importe2,texto5) "
    
'    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
'
'    RC = RC & "impvenci + if(gastos is null,0,gastos) - if( impcobro is null,0,impcobro) ,credicon"
'    RC = RC & "  from conta" & NumConta & ".scobro,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
'    RC = RC & " WHERE scobro.codmacta=cuentas.codmacta  and numpoliz<>''  and "
'    RC = RC & " (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
'
    
    'ENERO 2013.
    'Despues de hablar con BERNIA, en este listado salen
    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
    
    RC = RC & " totfaccl ,credicon,if(pais is null,'',pais)"    'JUL13 añadimos PAIS
    RC = RC & " from conta" & NumConta & ".cabfact,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
    RC = RC & " WHERE cabfact.codmacta=cuentas.codmacta  and numpoliz<>''  and "
    
    RC = RC & " (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
    
    'Contemplamos facturas desde la fecha de concesion
    RC = RC & " and fecfaccl>= fecconce"
    
    'D/H fecha factura
    If Me.Text3(34).Text <> "" Then RC = RC & " AND fecfaccl >='" & Format(Text3(34).Text, FormatoFecha) & "'"
    If Me.Text3(35).Text <> "" Then RC = RC & " AND fecfaccl <='" & Format(Text3(35).Text, FormatoFecha) & "'"
    
    
    
    
    
    
    Sql = Sql & RC
    Conn.Execute Sql
End Sub


Private Function GeneraDatosFrasAsegurados() As Boolean
Dim NumConta As Byte

    NumConta = CByte(vEmpresa.codempre)
    GeneraDatosFrasAsegurados = False

    Sql = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4,"
    Sql = Sql & " importe1,  importe2,fecha1,fecha2) "
    
    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
    
    RC = RC & "impvenci + if(gastos is null,0,gastos) - if( impcobro is null,0,impcobro) ,if (credicon is null,0,credicon)"
    RC = RC & ",fecfaccl,fecvenci"
    RC = RC & "  from conta" & NumConta & ".scobro,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
    RC = RC & " WHERE scobro.codmacta=cuentas.codmacta  "
    
    If Me.chkVarios(0).Value = 1 Then
        'SOLO asegudaros
        RC = RC & " and numpoliz<>''  and (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
    End If
    'D/H fecha factura
    If Me.Text3(34).Text <> "" Then RC = RC & " AND fecfaccl >='" & Format(Text3(34).Text, FormatoFecha) & "'"
    If Me.Text3(35).Text <> "" Then RC = RC & " AND fecfaccl <='" & Format(Text3(35).Text, FormatoFecha) & "'"
    
    
    Sql = Sql & RC
    Conn.Execute Sql

    
    
    'Borramos importe cero

    Sql = "DELETE from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Sql = Sql & " AND importe1<=0"
    Conn.Execute Sql
    
    Sql = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
    If Sql <> "" Then
        NumRegElim = Val(Sql)
    Else
        NumRegElim = 0
    End If
    GeneraDatosFrasAsegurados = NumRegElim > 0



End Function

'****************************************************************************************
'****************************************************************************************
'
'       NORMA 57
'
'****************************************************************************************
'****************************************************************************************
Private Function procesarficheronorma57() As Boolean
Dim Estado As Byte  '0  esperando cabcerea
                    '1  esperando pie (leyendo lineas)
    
    On Error GoTo eprocesarficheronorma57
    
    
    'insert into tmpconext(codusu,cta,fechaent,Pos)
    Conn.Execute "DELETE FROM tmpconext WHERE codusu = " & vUsu.Codigo
    procesarficheronorma57 = False
    I = FreeFile
    Open cd1.FileName For Input As #I
    Sql = ""
    Estado = 0
    Importe = 0
    TotalRegistros = 0
    While Not EOF(I)
            Line Input #I, Sql
            RC = Mid(Sql, 1, 4)
            Select Case Estado
            Case 0
                'Para saber que el fichero tiene el formato correcto
                If RC = "0270" Then
                        Estado = 1
                        'Voy a buscar si hay un banco
                        
                        RC = "select cuentas.codmacta,nommacta from ctabancaria,cuentas where ctabancaria.codmacta="
                        RC = RC & "cuentas.codmacta AND ctabancaria.entidad = " & Trim(Mid(Sql, 23, 4))
                        Set miRsAux = New ADODB.Recordset
                        miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        TotalRegistros = 0
                        While Not miRsAux.EOF
                            RC = miRsAux!codmacta & "|" & miRsAux!Nommacta & "|"
                            TotalRegistros = TotalRegistros + 1
                            miRsAux.MoveNext
                        Wend
                        miRsAux.Close
                        If TotalRegistros = 1 Then
                            Me.txtCtaBanc(5).Text = RecuperaValor(RC, 1)
                            Me.txtDescBanc(5).Text = RecuperaValor(RC, 2)
                        End If
                        TotalRegistros = 0
                End If
            Case 1
                If RC = "6070" Then
                    'Linea con recibo
                    'Ejemplo:
                    '   6070      46076147000130582263151014000000014067003059                      0000000516142
                    '                                  fecha       impot   socio                      fra      CC codigo de control del codigo de barra
                    'Fecha pago
                    RC = Mid(Sql, 31, 2) & "/" & Mid(Sql, 33, 2) & "/20" & Mid(Sql, 35, 2)
                    Fecha = CDate(RC)
                    'IMporte
                    RC = Mid(Sql, 37, 12)
                    cad = CStr(CCur(Val(RC) / 100))
                    'FRA
                    RC = Mid(Sql, 77, 11)
                    CONT = Val(RC)
                    'Socio
                    RC = Val(Mid(Sql, 50, 6))
                        
                    'Insertamos en tmp
                    TotalRegistros = TotalRegistros + 1
                    Sql = "INSERT INTO tmpconext(codusu,cta,fechaent,Pos,TimporteD,linliapu) VALUES (" & vUsu.Codigo & ",'"
                    Sql = Sql & RC & "','" & Format(Fecha, FormatoFecha) & "'," & CONT & "," & TransformaComasPuntos(cad) & "," & TotalRegistros & ")"
                    Conn.Execute Sql
                    
                    Importe = Importe + CCur(TransformaPuntosComas(cad))
                ElseIf RC = "8070" Then
                    'OK. Final de linea.
                    '
                    'Comprobacion BASICA
                    '8070      46076147000 000010        000000028440
                    '                       vtos-2           importe
                    
                    RC = ""
                    
                    'numero registros
                    cad = Val(Mid(Sql, 24, 5))
                    If Val(cad) = 0 Then
                        RC = RC = RC & vbCrLf & "Linea totales. Nº registros cero. " & Sql
                    Else
                        If Val(cad) - 2 <> TotalRegistros Then RC = "Contador de registros incorrecto"
                    End If
                    'Suma importes
                    cad = CStr(CCur(Mid(Sql, 37, 12) / 100))
                    
                    If CCur(cad) = 0 Then
                        RC = RC = RC & vbCrLf & "Linea totales. Suma importes cero. " & Sql
                    Else
                        If CCur(cad) <> Importe Then RC = RC & vbCrLf & "Suma importes incorrecta"
                    End If
                    
                    
                   
                    
                    
                    If RC <> "" Then
                        Err.Raise 513, , RC
                    Else
                        Estado = 2
                    End If
                End If
            End Select
    Wend
    Close #I
    I = 0 'para que no vuelva a cerrar el fichero
    
    If Estado < 2 Then
        'Errores procesando fichero
        If Estado = 0 Then
            Sql = "No se encuetra la linea de inicio de declarante(6070)"
        Else
            Sql = "No se encuetra la linea de totales(8070)"
        End If

        MsgBox "Error procesando el fichero." & vbCrLf & Sql, vbExclamation
    Else
        espera 0.5
        procesarficheronorma57 = True
    End If
eprocesarficheronorma57:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    If I > 0 Then Close #I
End Function


Private Function BuscarVtosNorma57() As Boolean

    BuscarVtosNorma57 = False
    
    Set miRsAux = New ADODB.Recordset
    
    'Dependiendo del parametro....
    If vParamT.Norma57 = 1 Then
        'ESCALONA.
        'Viene el socio y el numero de factura e importe.
        'Habra que buscar
        BuscarVtosNorma57 = VtosNorma57Escalona

    Else
        MsgBox "En desarrollo", vbExclamation
    End If
    
    Set miRsAux = Nothing
End Function

Private Function VtosNorma57Escalona() As Boolean
Dim RN As ADODB.Recordset
Dim Fin As Boolean
Dim NoEncontrado As Byte
Dim AlgunVtoNoEncontrado As Boolean
On Error GoTo eVtosNorma57Escalona
    
    VtosNorma57Escalona = False
    Set RN = New ADODB.Recordset
    Sql = "select * from tmpconext WHERE codusu =" & vUsu.Codigo & " order by cta,pos "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AlgunVtoNoEncontrado = False
    While Not miRsAux.EOF
        'Vto a vto
        'If miRsAux!Linliapu = 9 Then Stop
        RC = RellenaCodigoCuenta("430." & miRsAux!Cta)
        Sql = "Select * from scobro where codmacta = '" & RC & "' AND codfaccl =" & miRsAux!Pos & " and impvenci>0"
        RN.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CONT = 0
        If RN.EOF Then
            cad = "NO encontrado"
            NoEncontrado = 2
        Else
            'OK encontrado.
            Fin = False
            I = 0
            NoEncontrado = 1
            cad = ""
            
            While Not Fin
            
                I = I + 1
                
                Norma57VencimientoEncontradoEsCorrecto RN, Fin
                
                If Not Fin Then
                    RN.MoveNext
                    If RN.EOF Then Fin = True
                End If
            Wend
        End If
        RN.Close
        Sql = "UPDATE tmpconext SET "
        If CONT = 1 Then
            'OK este es el vto
            'NO hacemos nada. Updateamos los campos de la tmp
            'para buscar despues
            'numdiari numorden       numdocum=fecfaccl     ccost numserie
            Sql = Sql & " nomdocum ='" & Format(Fecha, FormatoFecha)
            Sql = Sql & "', ccost ='" & DevfrmCCtas
            Sql = Sql & "', numdiari = " & I
            Sql = Sql & ", contra = '" & RC & "'"
        Else
            If I > 1 Then cad = "(+1) " & cad
            Sql = Sql & " numasien=  " & NoEncontrado  'para vtos no encontrados o erroneos
            Sql = Sql & ", ampconce ='" & DevNombreSQL(cad) & "'"
            If NoEncontrado = 2 Then AlgunVtoNoEncontrado = True
        End If
        Sql = Sql & " WHERE codusu =" & vUsu.Codigo & " AND linliapu = " & miRsAux!Linliapu
        Conn.Execute Sql
            
 
        
        'Sig
        miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    
    
    
    If AlgunVtoNoEncontrado Then
        'Lo buscamos al reves
        espera 0.5
        Sql = "select * from  tmpconext  WHERE codusu =" & vUsu.Codigo & " AND numasien=2"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Miguel angel
            'Puede que en algunos recibos las posciones del fichero vengan cambiadas
            'Donde era la factura es la cta y al reves
            RC = RellenaCodigoCuenta("430." & miRsAux!Pos)
            Sql = "Select * from scobro where codmacta = '" & RC & "' AND codfaccl =" & Val(miRsAux!Cta) & " and impvenci>0"
            RN.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RN.EOF Then
        
                'OK encontrado.
                Fin = False
                CONT = 0
                Norma57VencimientoEncontradoEsCorrecto RN, Fin
                
                
            
            
                'OK este es el vto
                'NO hacemos nada. Updateamos los campos de la tmp
                'para buscar despues
                'numdiari numorden       numdocum=fecfaccl     ccost numserie
                If CONT = 1 Then
                    Sql = Sql & " nomdocum ='" & Format(Fecha, FormatoFecha)
                    Sql = Sql & "', ccost ='" & DevfrmCCtas
                    Sql = Sql & "', numdiari = " & I
                    Sql = Sql & ", contra = '" & RC & "'"
                    Sql = "UPDATE tmpconext SET "
                End If
            End If
            RN.Close
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    
    
    
    
    VtosNorma57Escalona = True
eVtosNorma57Escalona:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, "Buscando vtos Escalona"
    Set RN = Nothing
End Function


Private Sub Norma57VencimientoEncontradoEsCorrecto(ByRef Rss As ADODB.Recordset, ByRef Final As Boolean)
        
        'Ha encontrado el vencimiento. Falta ver si no esta en remesa....
        If Not IsNull(Rss!CodRem) Then
            cad = "En la remesa " & Rss!CodRem
        
        Else
            If Not IsNull(Rss!transfer) Then
                cad = "Transferencia " & Rss!transfer
            Else
                Importe = Rss!ImpVenci + DBLet(Rss!Gastos, "N") - DBLet(Rss!impcobro, "N")
                If Importe <> miRsAux!timported Then
                    'Importe distinto
                    'Veamos si es que esta
                    cad = "Importe distinto"
                Else
                    'OK. Misma factura, socio, importe. SAlimos ya poniendo ""
                    Fecha = Rss!fecfaccl
                    DevfrmCCtas = Rss!NUmSerie
                    I = Rss!numorden
                    cad = ""
                    Final = True
                    CONT = 1
                End If
            End If
        End If
End Sub

Private Sub CargaLWNorma57(Correctos As Boolean)
Dim IT As ListItem

    Set miRsAux = New ADODB.Recordset
    If Correctos Then
        Sql = "select tmpconext.*,nommacta from tmpconext left join cuentas on tmpconext.contra=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
        Sql = Sql & " and numasien=0 order by  ccost,pos  "
    Else
        Sql = "select * from tmpconext WHERE codusu = " & vUsu.Codigo & " and numasien > 0 order by cta,pos "
    End If
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Correctos Then
            Set IT = Me.lwNorma57Importar(0).ListItems.Add(, "C" & Format(miRsAux!Linliapu, "0000"))
            IT.Text = miRsAux!CCost
            IT.SubItems(1) = miRsAux!Pos
            IT.SubItems(2) = Format(miRsAux!Nomdocum, "dd/mm/yyyy")
            IT.SubItems(3) = miRsAux!Linliapu
            If IsNull(miRsAux!Nommacta) Then
                Sql = "ERRROR GRAVE"
            Else
                Sql = miRsAux!Nommacta
            End If
            IT.SubItems(4) = Sql
            IT.SubItems(5) = Format(miRsAux!timported, FormatoImporte)
            IT.SubItems(6) = Format(miRsAux!FechaEnt, "dd/mm/yyyy")
            IT.Checked = True
        Else
            'ERRORES
            Set IT = Me.lwNorma57Importar(1).ListItems.Add(, "C" & Format(miRsAux!Linliapu, "0000"))
            IT.Text = miRsAux!Cta
            IT.SubItems(1) = miRsAux!Pos
            IT.SubItems(2) = Format(miRsAux!timported, FormatoImporte)
            IT.SubItems(3) = miRsAux!Ampconce
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub


Private Sub LimpiarDelProceso()
    lwNorma57Importar(0).ListItems.Clear
    lwNorma57Importar(1).ListItems.Clear
    Me.txtCtaBanc(5).Text = ""
    Me.txtDescBanc(5).Text = ""
End Sub



