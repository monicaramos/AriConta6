VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelarAccion 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   5400
      TabIndex        =   147
      Top             =   6480
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameCtaExpCC 
      Height          =   5640
      Left            =   120
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   5715
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Solo mostrar subcentros de reparto"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   58
         Top             =   4320
         Width           =   2775
      End
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Comparativo"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   59
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   30
         Left            =   1320
         TabIndex        =   51
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   30
         Left            =   2520
         TabIndex        =   156
         Text            =   "Text5"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   29
         Left            =   1320
         TabIndex        =   50
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   2520
         TabIndex        =   154
         Text            =   "Text5"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Frame FrameCCComparativo 
         BorderStyle     =   0  'None
         Caption         =   "Frame12"
         Height          =   495
         Left            =   1920
         TabIndex        =   153
         Top             =   4560
         Visible         =   0   'False
         Width           =   3375
         Begin VB.OptionButton optCCComparativo 
            Caption         =   "Mes"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   61
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton optCCComparativo 
            Caption         =   "Saldo"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   60
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Ver movimientos posteriores"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   4320
         Width           =   2415
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   7
         ItemData        =   "frmListado.frx":030A
         Left            =   4200
         List            =   "frmListado.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   3
         Left            =   1380
         TabIndex        =   49
         Text            =   "Text2"
         Top             =   1440
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2400
         TabIndex        =   65
         Text            =   "Text2"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   2
         Left            =   1380
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   1020
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   64
         Text            =   "Text2"
         Top             =   1020
         Width           =   2655
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   4440
         TabIndex        =   63
         Top             =   5160
         Width           =   975
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   6
         ItemData        =   "frmListado.frx":030E
         Left            =   1620
         List            =   "frmListado.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   8
         Left            =   2940
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   3840
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   5
         ItemData        =   "frmListado.frx":0312
         Left            =   1620
         List            =   "frmListado.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   7
         Left            =   2940
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdCtaExpCC 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   62
         Top             =   5160
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   119
         Left            =   240
         TabIndex        =   158
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   73
         Left            =   480
         TabIndex        =   157
         Top             =   2685
         Width           =   465
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   30
         Left            =   1080
         Picture         =   "frmListado.frx":0316
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   480
         TabIndex        =   155
         Top             =   2325
         Width           =   465
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   29
         Left            =   1080
         Picture         =   "frmListado.frx":0D18
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   315
         Left            =   180
         TabIndex        =   74
         Top             =   5160
         Width           =   2835
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes cálculo"
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
         Left            =   4200
         TabIndex        =   73
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta explotación por centro de coste"
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
         Height          =   465
         Index           =   12
         Left            =   300
         TabIndex        =   72
         Top             =   300
         Width           =   5025
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de coste"
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
         Left            =   240
         TabIndex        =   71
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   35
         Left            =   540
         TabIndex        =   70
         Top             =   1140
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmListado.frx":171A
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   540
         TabIndex        =   69
         Top             =   1500
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmListado.frx":211C
         Top             =   1470
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   33
         Left            =   960
         TabIndex        =   68
         Top             =   3420
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   32
         Left            =   960
         TabIndex        =   67
         Top             =   3900
         Width           =   495
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
         Index           =   40
         Left            =   240
         TabIndex        =   66
         Top             =   3240
         Width           =   585
      End
   End
   Begin VB.Frame FrameBalancesper 
      Height          =   4395
      Left            =   30
      TabIndex        =   129
      Top             =   90
      Visible         =   0   'False
      Width           =   6075
      Begin VB.CheckBox chkApaisado 
         Caption         =   "Apaisado"
         Height          =   255
         Left            =   2880
         TabIndex        =   152
         Top             =   2520
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame FrameTapa2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   600
         TabIndex        =   146
         Top             =   2760
         Width           =   3795
      End
      Begin VB.CheckBox chkBalPerCompa 
         Caption         =   "Comparativo"
         Height          =   255
         Left            =   720
         TabIndex        =   144
         Top             =   2520
         Width           =   1515
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   25
         Left            =   180
         TabIndex        =   138
         Top             =   3900
         Width           =   1455
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   25
         Left            =   4800
         TabIndex        =   137
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdBalances 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   136
         Top             =   3780
         Width           =   975
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   17
         ItemData        =   "frmListado.frx":2B1E
         Left            =   1440
         List            =   "frmListado.frx":2B20
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   2940
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   17
         Left            =   2880
         TabIndex        =   134
         Text            =   "Text1"
         Top             =   2940
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   16
         ItemData        =   "frmListado.frx":2B22
         Left            =   1440
         List            =   "frmListado.frx":2B24
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   16
         Left            =   2880
         TabIndex        =   132
         Text            =   "Text1"
         Top             =   1980
         Width           =   855
      End
      Begin VB.TextBox txtNumBal 
         Height          =   315
         Index           =   0
         Left            =   780
         TabIndex        =   131
         Text            =   "Text1"
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox TextDescBalance 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   130
         Text            =   "Text1"
         Top             =   1140
         Width           =   4035
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   54
         Left            =   720
         TabIndex        =   145
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   25
         Left            =   1500
         Picture         =   "frmListado.frx":2B26
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Left            =   180
         TabIndex        =   143
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label Label17 
         Caption         =   "Balances configurables"
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
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   142
         Top             =   300
         Width           =   4875
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   55
         Left            =   720
         TabIndex        =   141
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   76
         Left            =   240
         TabIndex        =   140
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
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
         Left            =   180
         TabIndex        =   139
         Top             =   780
         Width           =   660
      End
      Begin VB.Image ImgNumBal 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "frmListado.frx":2BB1
         Top             =   1140
         Width           =   240
      End
   End
   Begin VB.Frame frameCCxCta 
      Height          =   6435
      Left            =   60
      TabIndex        =   75
      Top             =   20
      Visible         =   0   'False
      Width           =   5655
      Begin VB.OptionButton optCCxCta 
         Caption         =   "SIN cen. reparto"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   151
         Top             =   4560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optCCxCta 
         Caption         =   "Centros de reparto"
         Height          =   195
         Index           =   1
         Left            =   3720
         TabIndex        =   150
         Top             =   4260
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optCCxCta 
         Caption         =   "Todo"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   149
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CheckBox chkCC_Cta 
         Caption         =   "Ver movimientos posteriores"
         Height          =   195
         Left            =   2880
         TabIndex        =   103
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2520
         TabIndex        =   99
         Text            =   "Text5"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   2520
         TabIndex        =   98
         Text            =   "Text5"
         Top             =   1260
         Width           =   2535
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   15
         Left            =   1320
         TabIndex        =   77
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   14
         Left            =   1320
         TabIndex        =   76
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton cmdCCxCta 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   85
         Top             =   5820
         Width           =   1035
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   10
         Left            =   2640
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   4380
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   8
         ItemData        =   "frmListado.frx":35B3
         Left            =   1320
         List            =   "frmListado.frx":35B5
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   9
         Left            =   2640
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   3900
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   9
         ItemData        =   "frmListado.frx":35B7
         Left            =   1320
         List            =   "frmListado.frx":35B9
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   4380
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   4440
         TabIndex        =   86
         Top             =   5820
         Width           =   975
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   2280
         TabIndex        =   88
         Text            =   "Text2"
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   79
         Text            =   "Text2"
         Top             =   3000
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   2280
         TabIndex        =   87
         Text            =   "Text2"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   78
         Text            =   "Text2"
         Top             =   2520
         Width           =   795
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   10
         ItemData        =   "frmListado.frx":35BB
         Left            =   1380
         List            =   "frmListado.frx":35BD
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Opciones"
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
         Index           =   109
         Left            =   3720
         TabIndex        =   148
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   46
         Left            =   180
         TabIndex        =   102
         Top             =   900
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   41
         Left            =   480
         TabIndex        =   101
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   40
         Left            =   480
         TabIndex        =   100
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   15
         Left            =   1020
         Picture         =   "frmListado.frx":35BF
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   14
         Left            =   1020
         Picture         =   "frmListado.frx":3FC1
         Top             =   1260
         Width           =   240
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
         Index           =   45
         Left            =   180
         TabIndex        =   97
         Top             =   3600
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   39
         Left            =   660
         TabIndex        =   96
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   38
         Left            =   660
         TabIndex        =   95
         Top             =   3960
         Width           =   615
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   5
         Left            =   1020
         Picture         =   "frmListado.frx":49C3
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   37
         Left            =   480
         TabIndex        =   94
         Top             =   3120
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   4
         Left            =   1020
         Picture         =   "frmListado.frx":53C5
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   36
         Left            =   480
         TabIndex        =   93
         Top             =   2580
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de coste"
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
         Index           =   44
         Left            =   180
         TabIndex        =   92
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Centros de coste por cuenta"
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
         Index           =   13
         Left            =   780
         TabIndex        =   91
         Top             =   360
         Width           =   4185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes cálculo"
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
         Index           =   43
         Left            =   240
         TabIndex        =   90
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label15"
         Height          =   315
         Index           =   26
         Left            =   180
         TabIndex        =   89
         Top             =   5940
         Width           =   2835
      End
   End
   Begin VB.Frame frameccporcta 
      Height          =   5235
      Left            =   -60
      TabIndex        =   104
      Top             =   90
      Visible         =   0   'False
      Width           =   5415
      Begin MSComctlLib.ProgressBar pb7 
         Height          =   375
         Left            =   120
         TabIndex        =   127
         Top             =   4620
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.CommandButton cmdCtapoCC 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   121
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   19
         Left            =   4200
         TabIndex        =   126
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   6
         Left            =   1560
         TabIndex        =   119
         Text            =   "Text2"
         Top             =   3240
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   2640
         TabIndex        =   122
         Text            =   "Text2"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   7
         Left            =   1560
         TabIndex        =   120
         Text            =   "Text2"
         Top             =   3720
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   2640
         TabIndex        =   118
         Text            =   "Text2"
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   20
         Left            =   3720
         TabIndex        =   114
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   19
         Left            =   1440
         TabIndex        =   113
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2640
         TabIndex        =   108
         Text            =   "Text5"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2640
         TabIndex        =   107
         Text            =   "Text5"
         Top             =   1260
         Width           =   2535
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   17
         Left            =   1440
         TabIndex        =   106
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   105
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Label20"
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   128
         Top             =   4320
         Width           =   5055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de coste"
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
         Index           =   57
         Left            =   120
         TabIndex        =   125
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   660
         TabIndex        =   124
         Top             =   3300
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   7
         Left            =   1320
         Picture         =   "frmListado.frx":5DC7
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   660
         TabIndex        =   123
         Top             =   3735
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   6
         Left            =   1260
         Picture         =   "frmListado.frx":67C9
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   117
         Top             =   2445
         Width           =   495
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
         Index           =   56
         Left            =   180
         TabIndex        =   116
         Top             =   2220
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   115
         Top             =   2445
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   20
         Left            =   3480
         Picture         =   "frmListado.frx":71CB
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   19
         Left            =   1200
         Picture         =   "frmListado.frx":7256
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label19 
         Caption         =   "Detalle de explotación centro de coste"
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
         Left            =   240
         TabIndex        =   112
         Top             =   360
         Width           =   4875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         TabIndex        =   111
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   45
         Left            =   600
         TabIndex        =   110
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   44
         Left            =   600
         TabIndex        =   109
         Top             =   1305
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   17
         Left            =   1200
         Picture         =   "frmListado.frx":72E1
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   16
         Left            =   1200
         Picture         =   "frmListado.frx":7CE3
         Top             =   1260
         Width           =   240
      End
   End
   Begin VB.Frame frameCCostSaldos 
      Height          =   4755
      Left            =   60
      TabIndex        =   29
      Top             =   20
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdSaldosCC 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   37
         Top             =   3960
         Width           =   1035
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   6
         Left            =   2880
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   3060
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   4
         ItemData        =   "frmListado.frx":86E5
         Left            =   1440
         List            =   "frmListado.frx":86E7
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3060
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   5
         Left            =   2880
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   2580
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   3
         ItemData        =   "frmListado.frx":86E9
         Left            =   1440
         List            =   "frmListado.frx":86EB
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2580
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   4020
         TabIndex        =   38
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   43
         Text            =   "Text2"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         TabIndex        =   46
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   31
         Left            =   720
         TabIndex        =   45
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   30
         Left            =   720
         TabIndex        =   44
         Top             =   2580
         Width           =   615
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmListado.frx":86ED
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   29
         Left            =   300
         TabIndex        =   42
         Top             =   1740
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmListado.frx":90EF
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   300
         TabIndex        =   40
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de coste"
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
         TabIndex        =   39
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Saldos centros de coste"
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
         Index           =   11
         Left            =   960
         TabIndex        =   33
         Top             =   360
         Width           =   3525
      End
   End
   Begin VB.Frame FrameBalPresupues 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkQuitarApertura 
         Caption         =   "Quitar apertura"
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtMes 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   17
         Text            =   "Text4"
         Top             =   2880
         Width           =   915
      End
      Begin MSComctlLib.ProgressBar pb4 
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   5160
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdBalPre 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5040
         TabIndex        =   19
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2640
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2640
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   13
         Left            =   1440
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkPreMensual 
         Caption         =   "Mensual"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkPreAct 
         Caption         =   "Ejercicio siguiente"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Frame FrameNivelbalPresu 
         Height          =   1035
         Left            =   120
         TabIndex        =   1
         Top             =   3720
         Width           =   5865
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   11
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   10
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   9
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   8
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4560
            TabIndex        =   6
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   5
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   2400
            TabIndex        =   4
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "Último:  "
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nivel     "
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
            Left            =   120
            TabIndex        =   27
            Top             =   0
            Width           =   630
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Index           =   124
         Left            =   240
         TabIndex        =   159
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Opciones"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Balance presupuestario"
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
         Height          =   375
         Index           =   7
         Left            =   1440
         TabIndex        =   25
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   600
         TabIndex        =   23
         Top             =   1680
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   13
         Left            =   1200
         Picture         =   "frmListado.frx":9AF1
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   12
         Left            =   1200
         Picture         =   "frmListado.frx":A4F3
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   600
         TabIndex        =   22
         Top             =   1245
         Width           =   465
      End
   End
   Begin VB.Menu mnP1 
      Caption         =   "p1"
      Visible         =   0   'False
      Begin VB.Menu mnPrueba 
         Caption         =   "Prueba F1"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Byte
    '1 .- Listado consultas extractos, listado de MAYOR
    '2 .- Listado de cuentas
    '3 .- Listado de asientos
    '4 .- Totales cuenta concepto
    '5 .- balance de sumas y saldos
    '6 .- Reemision de diario
    '7 .- Cuentas de explotacion
    '8 .- Listado facturas clientes
    '9 .- Presupuestos
    
    '10 .- Balance presupuestario
    '11 .- Certificado declaración de IVA
    '12 .- Liquidacion IVA
    '13 .- Listado facturas proveedores    LO CAMBIAMOS 19/FEB/2004
    '14 .- Libro diario oficial
    
    '       Centros de coste
    '15 .- Acumulados y saldos
    '16 .- Cuenta explotacion centro de coste
    '17 .- Centro de coste por cuenta
        
    '18 .- Diario resumen
    '19 .- Cta explotacion por cta
    
    '20 .- Modelo 347
    '21 .- Cuenta de explotacion comparativa
    
    '22 .- Borre facturas clientes
    '23 .- "         "    proveedores
    '24 .- Balance consolidado de empresas
    
    '25 .- Balances personalizados
    '26 .-   "          "           Perdeterminado Situacion
    '27 .-   "          "           Perdeterminado Py g
    
    
    '28 .- Modelo 349
    '29 .- Traspaso PERSA
    '30 .- Traspaso ACE
    
    '31 .- Cuenta explotacion CONSIOLIDADA
    
    
    '----------------------------------- Legalizacion de libros
    ' 32.- Diario Normal. Como el 14
    ' 33.- Diario resumen. Como el 18
    ' 34.- Consulta extracots
    ' 35.- Inventario inicial
    ' 36.- Balance sumas y saldos
    ' 37.- Listado facturas clientes
    ' 38.- Listado Facturas proveedores
    ' 39.- Balance pyG
    ' 40.- Balance Situacion
    ' 41.- Inventario final
    
    ' Antiguos 39 y 40
    ' 50 .- Balance perosnalizados, consolidados. PyG
    ' 51 .- "           "               "        SITUACION
    
    ' 52 .- Facturas proveedor Consolidadas
    ' 53 .- Facturas CLIENTES   consolidada

    ' 54 .- Evolucion mensual de saldos

    ' 55 .- Relacion de clientes por cuenta gastos/ventas
    ' 56 .-  "          proveedores   ""           ""

    ' 57 .- Copiar balance configurables
    ' 58 .- Modelo 340
    
Public EjerciciosCerrados As Boolean
    'En algunos informes me servira para utilizar unas tablas u otras
Public Legalizacion As String   'Datos para la legalizacion
    
Dim Tablas As String
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCoste
Attribute frmCC.VB_VarHelpID = -1

Dim SQL As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim Cont As Long
Dim i As Integer

Dim Importe As Currency

'Para los balcenes frameBalance
' Cuando este trbajando con cerrado
' Para poder sbaer cuando empezaba el año del ejercicio a listar
Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date


Dim HanPulsadoSalir As Boolean

'Para cancelar
Dim PulsadoCancelar As Boolean


Private Sub chkApaisado_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub chkBalPerCompa_Click()
    FrameTapa2.Visible = Me.chkBalPerCompa.Value = 0
End Sub


Private Sub chkBalPerCompaCon_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkCC_Cta_KeyPress(KeyAscii As Integer)
        ListadoKEYpress KeyAscii
End Sub


Private Sub ChkConso_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkCtaExpCC_Click(Index As Integer)
    If Index = 1 Then
         FrameCCComparativo.Visible = chkCtaExpCC(1).Value = 1
    End If
End Sub

Private Sub chkCtaExpCC_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub



Private Sub ChkCtaPre_Click(Index As Integer)
'    If ChkCtaPre(Index).Value = 1 Then
'        For I = 1 To 10
'            If I <> Index Then ChkCtaPre(I).Value = 0
'        Next I
'    End If
End Sub

Private Sub chkDesgloseEmpresa_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkPreAct_KeyPress(KeyAscii As Integer)
ListadoKEYpress KeyAscii
End Sub

Private Sub chkPreMensual_KeyPress(KeyAscii As Integer)
ListadoKEYpress KeyAscii
End Sub

Private Sub cmbFecha_Click(Index As Integer)
    If Not PrimeraVez Then
        If Index = 0 Then ComprobarFechasBalanceQuitar6y7
    End If
End Sub

Private Sub cmbFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub cmbFecha_KeyPress(Index As Integer, KeyAscii As Integer)

    ListadoKEYpress KeyAscii
    
End Sub







Private Sub cmdBalances_Click()
    'Comprobamos datos
    If Me.txtNumBal(0).Text = "" Then
        MsgBox "Número de balance incorrecto", vbExclamation
        Exit Sub
    End If
    
    
    'Año 1
    If txtAno(16).Text = "" Then
        MsgBox "Año no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    If Val(txtAno(16).Text) < 1900 Then
        MsgBox "No se permiten años anteriores a 1900", vbExclamation
        Exit Sub
    End If
    
    If chkBalPerCompa.Value = 1 Then
        If txtAno(17).Text = "" Then
            MsgBox "Año no puede estar en blanco", vbExclamation
            Exit Sub
        End If
        If Val(txtAno(17).Text) < 1900 Then
            MsgBox "No se permiten años anteriores a 1900", vbExclamation
            Exit Sub
        End If
    End If

    'Fecha informe
    If Text3(25).Text = "" Then
        MsgBox "Fecha informe incorrecta.", vbExclamation
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    i = -1
    If chkBalPerCompa.Value = 1 Then
        i = Val(cmbFecha(17).ListIndex)
        i = i + 1
        If i = 0 Then i = -1
    End If
    GeneraDatosBalanceConfigurable CInt(txtNumBal(0).Text), Me.cmbFecha(16).ListIndex + 1, CInt(txtAno(16).Text), i, Val(txtAno(17).Text), False, -1
    
    
    'Para saber k informe abriresmos
    Cont = 1
    RC = 1 'Perdidas y ganancias
    Set Rs = New ADODB.Recordset
    SQL = "Select * from sbalan where numbalan=" & Me.txtNumBal(0).Text
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then

            If DBLet(Rs!Aparece, "N") = 0 Then
                Cont = 3
            Else
                Cont = 1
            End If

        RC = Rs!perdidas
    End If
    Rs.Close
    Set Rs = Nothing
        
        
    'Si es comarativo o no
    If Me.chkBalPerCompa.Value = 1 Then Cont = Cont + 1
        
    'Textos
    RC = "perdidasyganancias= " & RC & "|"
          
    SQL = RC & "FechaImp= """ & Text3(25).Text & """|"
    SQL = SQL & "Titulo= """ & Me.TextDescBalance(0).Text & """|"
    'PGC 2008 SOlo pone el año, NO el mes
    If vParam.NuevoPlanContable Then
        RC = ""
    Else
        RC = cmbFecha(16).List(cmbFecha(16).ListIndex)
    End If
    RC = RC & " " & txtAno(16).Text
    RC = "fec1= """ & RC & """|"
    SQL = SQL & RC
    
    If Me.chkBalPerCompa.Value = 1 Then
            'PGC 2008 SOlo pone el año, NO el mes
            If vParam.NuevoPlanContable Then
                RC = ""
            Else
                RC = cmbFecha(17).List(cmbFecha(17).ListIndex)
            End If
            RC = RC & " " & txtAno(17).Text
            RC = "Fec2= """ & RC & """|"
            SQL = SQL & RC
            

    Else
        'Pong el nombre del mes
        RC = UCase(Mid(cmbFecha(16).Text, 1, 1)) & Mid(cmbFecha(16).Text, 2, 2)
        RC = "vMes= """ & RC & """|"
        SQL = SQL & RC
    End If
    SQL = SQL & "Titulo= """ & Me.TextDescBalance(0).Text & """|"
        
    If opcion < 39 Or opcion > 40 Then
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 4
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            
            'La opcion sera si esta marcado apaisado
            If chkApaisado.Value = 1 Then
                .opcion = 82 + Cont   'El 83 es el primero en la de apisado que es para el PGC2008
            Else
                .opcion = 48 + Cont   'El 49 es el primero de los rpt de balance
            End If
            .Show vbModal
        End With
    Else
        GeneraLegalizaPRF SQL, 6
        CadenaDesdeOtroForm = "OK"
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdBalPre_Click()

    If Not ComprobarCuentas(12, 13) Then Exit Sub
    
    SQL = ""
    For i = 1 To Me.ChkCtaPre.Count
        If Me.ChkCtaPre(i).Value Then SQL = SQL & "&"
    Next i
    If Len(SQL) <> 1 Then
        If chkPreMensual.Value = 1 Then
            MsgBox "Seleccione uno, y solo uno, de los niveles contables.", vbExclamation
            Exit Sub
        End If
    End If
    
    
    If txtMes(2).Text <> "" And Me.chkPreMensual.Value = 0 Then
        
        MsgBox "Si indica el mes debe marcar la opcion ""mensual""", vbExclamation
        Exit Sub
    End If
    
    If txtMes(2).Text <> "" Then
        If Val(txtMes(2).Text) < 1 Or Val(txtMes(2).Text) > 12 Then
            MsgBox "Mes incorrecto: " & txtMes(2).Text, vbExclamation
            Exit Sub
        End If
    End If
    
    
    'Solo podemos quitar el asiento de apertura para ejercicio actual
    i = 0
    If chkQuitarApertura.Value = 1 Then
        i = 1
        'ejer siguiente
        If chkPreAct.Value = 1 Then
            i = 0
        Else
            'Si es mensual y el mes NO es uno tampoco lo quita
            If chkPreMensual.Value = 1 Then
                If Val(txtMes(2).Text) > 1 Then i = 0
            End If
        End If
    End If
    chkQuitarApertura.Value = i
        
    
    
    
    SQL = ""
    RC = ""
    If txtCta(12).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        RC = "Desde " & txtCta(12).Text & " - " & DtxtCta(12).Text
        SQL = SQL & "presupuestos.codmacta >= '" & txtCta(12).Text & "'"
    End If
    
    
    If txtCta(13).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        If RC <> "" Then
            RC = RC & "       h"
        Else
            RC = "H"
        End If
        RC = RC & "asta " & txtCta(13).Text & " - " & DtxtCta(13).Text
        SQL = SQL & "presupuestos.codmacta <= '" & txtCta(13).Text & "'"
    End If

    If SQL <> "" Then SQL = SQL & " AND"
    i = Year(vParam.fechaini)
    If chkPreAct.Value Then i = i + 1
    SQL = SQL & " anopresu =" & i
    
    
    If RC <> "" Then RC = """ + chr(13) +""" & RC
    If chkPreMensual.Value = 1 Then
        If txtMes(2).Text <> "" Then RC = "** " & Format("01/" & txtMes(2).Text & "/1999", "mmmm") & " ** " & RC
        RC = "  MENSUAL " & RC
    End If
    
    
    
    RC = "Año: " & i & RC
    CadenaDesdeOtroForm = ""
    
    For Cont = 1 To 10
        If ChkCtaPre(Cont).Value = 1 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "- " & Cont
    Next

    RC = RC & " Digitos: " & Mid(CadenaDesdeOtroForm, 2)
    
    If chkQuitarApertura.Value = 1 Then RC = RC & "     Sin apertura"
    CadenaDesdeOtroForm = "CampoSeleccion= """ & RC & """|"

    RC = ""
    For Cont = 1 To 9
        If ChkCtaPre(Cont).Value = 1 Then
            If RC = "" Then RC = Cont
        End If
    Next
    If RC = "" Then RC = "11"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Remarcar= " & RC & "|"
    


    If GeneraBalancePresupuestario() Then
        With frmImprimir
            .OtrosParametros = CadenaDesdeOtroForm
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .opcion = 23 + Me.chkPreMensual.Value
            .Show vbModal
        End With
    End If
    pb4.Visible = False
End Sub


Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub cmdCanListExtr_Click(Index As Integer)
    If Me.cmdCancelarAccion.Visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub

Private Sub cmdCCxCta_Click()

    '// Centros de coste por cuenta de explotacion

    If txtCCost(4).Text <> "" And txtCCost(5).Text <> "" Then
        If txtCCost(5).Text > txtCCost(5).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtAno(9).Text = "" Or txtAno(10).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Sub
    End If
    
    If txtAno(9).Text <> "" And txtAno(10).Text <> "" Then
        If Val(txtAno(9).Text) > Val(txtAno(10).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Sub
        Else
            If Val(txtAno(9).Text) = Val(txtAno(10).Text) Then
                If Me.cmbFecha(8).ListIndex > Me.cmbFecha(9).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    End If

    
    If Me.cmbFecha(10).ListIndex < 0 Then
        MsgBox "Seleccione un mes de cálculo", vbExclamation
        Exit Sub
    End If
    
    
    'Comprobamos que el total de meses no supera el año
    i = Val(txtAno(9).Text)
    Cont = Val(txtAno(10).Text)
    Cont = Cont - i
    i = 0
    If Cont > 1 Then
       i = 1  'Ponemos a uno para luego salir del bucle
    Else
        If Cont = 1 Then
            'Se diferencian un año, luego el mes fin tienes k ser menor a mes inicio
            If Me.cmbFecha(9).ListIndex >= Me.cmbFecha(8).ListIndex Then i = 1
        End If
    End If
    If i <> 0 Then
        MsgBox "El intervalo tiene que ser de un año como máximo", vbExclamation
        Exit Sub
    End If
    
    
    
    Screen.MousePointer = vbHourglass
    If GeneraCCxCtaExplotacion Then
        
        Label2(26).Caption = ""
        'Vamos a poner los textos
        SQL = "Mes cálculo: " & UCase(cmbFecha(10).List(cmbFecha(10).ListIndex))
        SQL = SQL & "   Desde : " & cmbFecha(8).ListIndex + 1 & " / " & txtAno(9).Text
        SQL = SQL & "   Hasta : " & cmbFecha(9).ListIndex + 1 & " / " & txtAno(10).Text
        
        
        Cad = ""
        If txtCta(14).Text <> "" Then Cad = "Desde cta:" & txtCta(14).Text
        If txtCta(15).Text <> "" Then
            If Cad <> "" Then Cad = Cad & "    "
            Cad = Cad & "Hasta cta: " & txtCta(15).Text
        End If
        If Cad <> "" Then SQL = SQL & "  " & Cad
        
        
        
        
        RC = ""
        'Centros de coste
        If Me.txtCCost(4).Text <> "" Then _
            RC = "Desde CC: " & Me.txtCCost(4).Text & " - " & Me.txtDCost(4).Text
        If Me.txtCCost(5).Text <> "" Then
            If RC <> "" Then RC = RC & "  "
            RC = RC & "Hasta CC: " & Me.txtCCost(5).Text & " - " & Me.txtDCost(5).Text
        End If
        
        
        If Me.chkCC_Cta.Value = 1 Then
            'Solo hay una linea
            i = 0
            Cont = 1
            If RC <> "" Then SQL = SQL & "     " & RC
            RC = ""
        Else
            'Hay dos lineas para poner todo
            i = 1
            Cont = 0
        End If
        
        
        'Si se ha marcado solo repartos lo marco en el informe
        If Me.optCCxCta(1).Value Then SQL = SQL & "   (C. REPARTO)"
        RC = "Fechas= """ & RC & """|"
        SQL = "Cuenta= """ & SQL & """|"
        SQL = SQL & RC
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = i + 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .opcion = 36 + Cont
            .Show vbModal
        End With
    End If
    Label15.Caption = ""
    Label2(26).Caption = ""
    Screen.MousePointer = vbDefault
    
    
End Sub




Private Sub cmdCtaExpCC_Click()

    If txtCCost(2).Text <> "" And txtCCost(3).Text <> "" Then
        If txtCCost(2).Text > txtCCost(3).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtAno(7).Text = "" Or txtAno(8).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Sub
    End If
    
    If Me.cmbFecha(7).ListIndex < 0 Then
        MsgBox "Seleccione un mes de cálculo", vbExclamation
        Exit Sub
    End If
    
    If Not ComparaFechasCombos(7, 8, 5, 6) Then Exit Sub
     
    
    'Comprobamos que el total de meses no supera el año
    i = Val(txtAno(7).Text)
    Cont = Val(txtAno(8).Text)
    Cont = Cont - i
    i = 0
    If Cont > 1 Then
       i = 1  'Ponemos a uno para luego salir del bucle
    Else
        If Cont = 1 Then
            'Se diferencian un año, luego el mes fin tienes k ser menor a mes inicio
            If Me.cmbFecha(6).ListIndex >= Me.cmbFecha(5).ListIndex Then i = 1
        End If
    End If
    If i <> 0 Then
        MsgBox "El intervalo tiene que ser de un año como máximo", vbExclamation
        Exit Sub
    End If


    'No puede pedir movimientos posteriores y comparativo
    If chkCtaExpCC(0).Value = 1 And chkCtaExpCC(1).Value = 1 Then
        MsgBox "No puede pedir comparativo y movimientos posteriores", vbExclamation
        Exit Sub
    End If


    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    Set Rs = New ADODB.Recordset
    If GeneraCtaExplotacionCC Then
        Label15.Caption = ""
        'Vamos a poner los textos
        If chkCtaExpCC(1).Value = 1 And optCCComparativo(0).Value Then
            SQL = ""
        Else
            SQL = "Mes cálculo: " & UCase(cmbFecha(7).List(cmbFecha(7).ListIndex)) & " "
        End If
        
        If chkCtaExpCC(2).Value = 1 Then SQL = Trim(SQL & "   Solo reparto") & "  "
        
        'If Not (chkCCComparativo.Value = 1 And optCCComparativo(1).Value) Then
        SQL = SQL & "Desde : " & cmbFecha(5).ListIndex + 1 & " / " & txtAno(7).Text
        SQL = SQL & " hasta : " & cmbFecha(6).ListIndex + 1 & " / " & txtAno(8).Text
        
        
        Cad = ""
        'Si han puesto desde hasta cuenta
        If txtCta(29).Text <> "" Then Cad = " Desde cta: " & txtCta(29).Text ' & " " & Mid(DtxtCta(29).Text, 1, 13) & "..."
        If txtCta(30).Text <> "" Then Cad = Cad & " hasta cta: " & txtCta(30).Text ' & " " & Mid(DtxtCta(30).Text, 1, 13) & "..."
        Cad = Trim(Cad)
    
    

  
        
        If Me.chkCtaExpCC(0).Value = 1 Then
            'Solo hay una linea
            RC = ""
            i = 0
            If Me.txtCCost(2).Text <> "" Then _
                SQL = SQL & "Desde CC: " & Me.txtCCost(2).Text & " - " & Me.txtDCost(2).Text
            If Me.txtCCost(3).Text <> "" Then _
                SQL = SQL & " Hasta CC: " & Me.txtCCost(3).Text & " - " & Me.txtDCost(3).Text
                
                
            'Cont = 1
            Cont = 35
        Else
        
                'Hay dos lineas para poner todo
                i = 1
                RC = ""
                If Me.txtCCost(2).Text <> "" Then _
                    RC = " Desde CC: " & Me.txtCCost(2).Text & " - " & Me.txtDCost(2).Text
                If Me.txtCCost(3).Text <> "" Then _
                    RC = RC & " Hasta CC: " & Me.txtCCost(3).Text & " - " & Me.txtDCost(3).Text
                'Cont = 0
                Cont = 34
                If chkCtaExpCC(1).Value = 1 Then   '2013  Octubre 28.  Habia chkCtaExpCC(0)
                    'Comparativo
                    Cont = 90
                    If optCCComparativo(0).Value Then Cont = Cont + 1
                End If

        End If
        SQL = Trim(SQL & "    " & Cad)
        RC = "Fechas= """ & RC & """|"
        SQL = "Cuenta= """ & SQL & """|"
        SQL = SQL & RC
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = i + 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            '.Opcion = 34 + Cont
            .opcion = Cont
            .Show vbModal
        End With
    End If
    Label15.Caption = ""
    Set miRsAux = Nothing
    Set Rs = Nothing
    Screen.MousePointer = vbDefault


End Sub







Private Sub cmdCtapoCC_Click()
Dim F As Date
'    If txtCta(16).Text <> "" And txtCta(17).Text <> "" Then
'        If Val(txtCta(16).Text) > Val(txtCta(17).Text) Then
'            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarCuentas(16, 17) Then Exit Sub

    If txtCCost(6).Text <> "" And txtCCost(7).Text <> "" Then
        If txtCCost(6).Text > txtCCost(7).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Sub
        End If
    End If
    If Not (Text3(19).Text <> "" And Text3(20).Text <> "") Then
        MsgBox "Debe introducir las fechas.", vbExclamation
        Exit Sub
    End If
'
'    If Text3(19).Text <> "" And Text3(20).Text <> "" Then
'        If CDate(Text3(19).Text) > CDate(Text3(20).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(19, 20) Then Exit Sub
    
    '-------------------------------------------------
    'INtervalo coja un año
    'Veamos siocupa mas de un año
    If Abs(DateDiff("d", CDate(Text3(19).Text), CDate(Text3(20).Text))) > 365 Then
        MsgBox "Las fechas son incorrectas. Abarca mas de un ejercicio", vbExclamation
        Exit Sub
    End If
    
    
    'Vamos a ver si coje un mismo año contable
    'Para ello situamos las fechas de inicio y fin de ejercicio
    'en funcion de la primera fecha
    F = CDate(Text3(19).Text)
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Años naturales
        FechaIncioEjercicio = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & Year(F))
        FechaFinEjercicio = CDate(Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & Year(F))
        Else
            'Años partidos
            'vemos si la donde entra le fecha de inicio
            'Auxiliarmente usamos este var
            FechaFinEjercicio = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & Year(F))
            i = Year(F)
                            'Es del años siguiente
            If F > FechaFinEjercicio Then i = i + 1
            
            'Ahora fijamos la de fin de jercicio y inicio
            FechaIncioEjercicio = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & i)
            FechaFinEjercicio = CDate(Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & i + 1)
    End If
    
    'Como era en funcion de la fecha de incio, comprobaremos la fecha fin
    F = CDate(Text3(20).Text)
    If F > FechaFinEjercicio Then
        MsgBox "Las fechas no estan dentro del mismo ejercicio contable", vbExclamation
        Exit Sub
    End If
    
    
    'Vemos si trabajamos con ejercicios cerrados
    F = UltimaFechaHcoCabapu
    If F >= FechaIncioEjercicio Then
        EjerciciosCerrados = True
    Else
        EjerciciosCerrados = False
    End If
    
    Screen.MousePointer = vbHourglass
    PulsadoCancelar = False
    Me.cmdCancelarAccion.Visible = True
    If ObtenerDatosCCCtaExp Then
        'Las cadenas
        SQL = "Desde " & Text3(19).Text & "  hasta  " & Text3(20).Text
            
        RC = ""
        If txtCta(16).Text <> "" Then RC = "Desde cuenta: " & txtCta(16).Text
        If txtCta(17).Text <> "" Then
        If RC = "" Then
                RC = "H"
            Else
                RC = RC & "  h"
            End If
            RC = RC & "asta cuenta: " & txtCta(17).Text
        End If
        
        
        If txtCCost(6).Text <> "" Then
            If RC <> "" Then RC = RC & "     "
            RC = RC & "Desde Centro coste: " & txtCCost(6).Text
        End If
                
        
        If txtCCost(7).Text <> "" Then
            If RC <> "" Then RC = RC & "     "
            RC = RC & "Hasta Centro coste: " & txtCCost(7).Text
        End If
        
        SQL = """" & SQL & """"
        If RC <> "" Then
            RC = " """ & RC & """"
            SQL = SQL & RC
        End If
        RC = "Cuenta= " & SQL & "|"
        
        With frmImprimir
            .OtrosParametros = RC
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .opcion = 41
            .Show vbModal
        End With
    End If
    Me.cmdCancelarAccion.Visible = False
    Label2(27).Visible = False
    pb7.Visible = False
    Screen.MousePointer = vbDefault
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



Private Sub cmdSaldosCC_Click()
Dim MesFin1 As Integer
Dim AnoFin1 As Integer

    On Error GoTo ESaldosCC
    If txtCCost(0).Text <> "" And txtCCost(1).Text <> "" Then
        If txtCCost(0).Text > txtCCost(1).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtAno(5).Text = "" Or txtAno(6).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Sub
    End If
    If Not ComparaFechasCombos(5, 6, 3, 4) Then Exit Sub
'    If txtAno(5).Text <> "" And txtAno(6).Text <> "" Then
'        If Val(txtAno(5).Text) > Val(txtAno(6).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        Else
'            If Val(txtAno(5).Text) = Val(txtAno(6).Text) Then
'                If Me.cmbFecha(4).ListIndex > Me.cmbFecha(4).ListIndex Then
'                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
    
    
    'Llegamos aqui y hacemos el sql, Para ello, y por si acaso piden de cerrados
    'tenemos k comprobar cual es el ultimo mes en saldosanal1
    UltimoMesAnyoAnal1 MesFin1, AnoFin1
    
    
    'Si años consulta iguales
    If txtAno(5).Text = txtAno(6).Text Then
         Cad = " anoccost=" & txtAno(5).Text & " AND mesccost>=" & Me.cmbFecha(3).ListIndex + 1
         Cad = Cad & " AND mesccost<=" & Me.cmbFecha(4).ListIndex + 1
         
    Else
        'Años disitintos
        'Inicio
        Cad = "( anoccost=" & txtAno(5).Text & " AND mesccost>=" & Me.cmbFecha(3).ListIndex + 1 & ")"
        Cad = Cad & " OR ( anoccost=" & txtAno(6).Text & " AND mesccost<=" & Me.cmbFecha(4).ListIndex + 1 & ")"
        'Por si la diferencia es mas de un año
        If Val(txtAno(6).Text) - Val(txtAno(5).Text) > 1 Then
            Cad = Cad & " OR (anoccost >" & txtAno(5).Text & " AND anoccost < " & txtAno(6).Text & ")"
        End If
    End If
    Cad = " AND (" & Cad & ")"
    
    RC = ""
    If txtCCost(0).Text <> "" Then RC = " cabccost.codccost >='" & txtCCost(0).Text & "'"
    If txtCCost(1).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " cabccost.codccost <='" & txtCCost(1).Text & "'"
    End If
    
    
    'Borramos temporal
    Screen.MousePointer = vbHourglass
    Conn.Execute "Delete from Usuarios.zsaldoscc  where codusu = " & vUsu.Codigo
    
    
    'Haremos las inserciones
    SQL = "INSERT INTO Usuarios.zsaldoscc (codusu, codccost, nomccost, ano, mes, impmesde, impmesha) SELECT "
    SQL = SQL & vUsu.Codigo & ",cabccost.codccost,nomccost,anoccost,mesccost,sum(debccost),sum(habccost) from hsaldosanal,cabccost where"
    SQL = SQL & " cabccost.codccost =hsaldosanal.codccost "
    If RC <> "" Then SQL = SQL & " AND " & RC
    SQL = SQL & Cad
    SQL = SQL & " group by codccost,anoccost,mesccost"
    Conn.Execute SQL
    
    
    
    
    'Haremos las inserciones desde hsaldosanal 1, es decir, ejercicios traspasados
    'si la fecha de incio de los calculos es  menor k la ultima fecha k haya en hco 1
    ' EN i tneemos el año y en mesfin1 el ultimo mes grabado en saldosanal1
    Tablas = ""
    If Val(txtAno(5).Text) < AnoFin1 Then
        Tablas = "SI"
    Else
        If Val(txtAno(5).Text) = AnoFin1 Then
            'Dependera del mes
            If MesFin1 >= (Me.cmbFecha(4).ListIndex + 1) Then Tablas = "OK"
        End If
    End If
    
    If Tablas <> "" Then
        SQL = "INSERT INTO Usuarios.zsaldoscc (codusu, codccost, nomccost, ano, mes, impmesde, impmesha) SELECT "
        SQL = SQL & vUsu.Codigo & ",cabccost.codccost,nomccost,anoccost,mesccost,sum(debccost),sum(habccost) from hsaldosanal1,cabccost where"
        SQL = SQL & " cabccost.codccost =hsaldosanal1.codccost "
        If RC <> "" Then SQL = SQL & " AND " & RC
        SQL = SQL & Cad
        SQL = SQL & " group by codccost,anoccost,mesccost"
        Conn.Execute SQL
    End If
    
    Set miRsAux = New ADODB.Recordset
    SQL = "Select count(mes) from Usuarios.zsaldoscc  where codusu = " & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 0
    If Not miRsAux.EOF Then
        i = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    
    If i = 0 Then
        MsgBox "Ningun dato con esos valores", vbExclamation
    Else
        'Titulitos
        RC = ""
        If txtCCost(0).Text <> "" Then RC = "Desde " & txtCCost(0).Text & " - " & txtDCost(0).Text
        If txtCCost(1).Text <> "" Then
            If RC = "" Then
                RC = "H"
            Else
                RC = RC & "     h"
            End If
            RC = RC & "asta " & txtCCost(1).Text & " - " & txtDCost(1).Text
        End If
        
        Cont = cmbFecha(3).ListIndex
        SQL = "Desde " & cmbFecha(3).List(Cont) & " - " & txtAno(5).Text
        Cont = cmbFecha(4).ListIndex
        SQL = SQL & "     hasta " & cmbFecha(4).List(Cont) & " - " & txtAno(6).Text
        
        If RC = "" Then
             RC = SQL
             SQL = ""
        End If
        Cad = "Cuenta= """ & RC & """|"
        Cad = Cad & "Fechas= """ & SQL & """|"
      
      
        
        
        With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .opcion = 33
            .Show vbModal
        End With

        
    End If
    
    
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ESaldosCC:
    MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Sub

    






Private Sub Form_Activate()

    

    If PrimeraVez Then
        PrimeraVez = False
        CommitConexion
        'Ponemos el foco
        Select Case opcion
        Case 1
            'Listado de EXTRACTOS DE cuentas
            '
        Case 2
            txtCta(3).SetFocus
        Case 3
    
        Case 4
            txtCta(5).SetFocus
        Case 5
            txtCta(6).SetFocus
        Case 6
        Case 7
            Text3(9).SetFocus
            
            
            
        Case 8
        Case 9
            txtCta(10).SetFocus
        Case 10
            txtCta(12).SetFocus
        Case 11
            Text3(12).SetFocus
        Case 12
        Case 13
            'txtNumFac(0).SetFocus
        Case 14
            Text3(15).SetFocus
        Case 15
            txtCCost(0).SetFocus
        Case 16
            txtCCost(2).SetFocus
        Case 17
            txtCta(14).SetFocus
        Case 18
            cmbFecha(11).SetFocus
        Case 21
            cmbFecha(13).SetFocus
            
        'Legalizacion de libros
        Case 32 To 41
            LegalizacionSub
                
        Case 52
            Text3(29).SetFocus
        Case 53
            Text3(10).SetFocus
            
        Case 54
            txtCta(23).SetFocus
        End Select
    End If
        Screen.MousePointer = vbDefault
End Sub

Private Sub LegalizacionSub()
            
            Screen.MousePointer = vbHourglass
            espera 0.1
            Me.Refresh
            Me.MousePointer = vbHourglass
            Select Case opcion
            Case 32
            Case 33
            Case 34
            Case 35
            Case 36, 41
            Case 37
            Case 38
            Case 39, 40
                cmdBalances_Click
            End Select
            Me.Hide
            espera 0.1
            Me.MousePointer = vbHourglass
            Unload Me


End Sub



Private Sub Form_Load()
Dim H As Single
Dim W As Single
    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    Limpiar Me
    
    'He puesto FALSE a todos los frames en diseño
    
'    FrameCuentas.Visible = False
'    frameListadoCuentas.Visible = False
'    frameAsiento.Visible = False
'    frameCtaConcepto.Visible = False
'    frameDiarioHco.Visible = False
'    frameBalance.Visible = False
'    frameExplotacion.Visible = False
'    frameListFacCli.Visible = False
'    FrameListFactP.Visible = False
'    Me.FramePresu.Visible = False
'    FrameBalPresupues.Visible = False
'    frameIVA.Visible = False
'    Frame4.Visible = False
'    Me.FrameLiq.Visible = False
'    Me.frameLibroDiario.Visible = False
'    frameCCostSaldos.Visible = False
'    frameCtaExpCC.Visible = False
'    frameCCxCta.Visible = False
'    frameResumen.Visible = False
'    frameccporcta.Visible = False
'    Frame347.Visible = False
'    Frame349.Visible = False
'    frameComparativo.Visible = False
'    frameBorrarClientes.Visible = False
'    Me.frameConsolidado.Visible = False
'    FrameBalancesper.Visible = False
'    FramePersa.Visible = False
'    FrameAce.Visible = False
'    frameExploCon.Visible = False
'    FrameBalPersoConso.Visible = False
    
    Select Case opcion
    Case 1, 34                  '34: Legalizacion
    Case 2
        'Listado de cuentas
    Case 3
        'Listado de asientos
    Case 4
    Case 5, 36, 41  '36: Legalizacion Bal sumas.
    Case 6, 35     'LEgalizacion Libros. Inventario Incial
    Case 7
    Case 8, 37, 53    '37: Presenacion telematica
    Case 9
    Case 10
        pb4.Visible = False
        PonerNiveles
        Me.FrameBalPresupues.Visible = True
        W = FrameBalPresupues.Width
        H = FrameBalPresupues.Height
        
    Case 11
    Case 12
        'Liquidacion IVA
    Case 13, 38, 52
        
    Case 14, 32   'El 32 es la impresion para el modelo de legaliza libros
        'Diario oficial
    Case 15
        frameCCostSaldos.Visible = True
        W = frameCCostSaldos.Width
        H = frameCCostSaldos.Height
        QueCombosFechaCargar "3|4|"
        cmbFecha(3).ListIndex = Month(vParam.fechaini) - 1
        cmbFecha(4).ListIndex = Month(vParam.fechafin) - 1
        txtAno(5).Text = Year(vParam.fechaini)
        txtAno(6).Text = Year(vParam.fechafin)
        
        
    Case 16
        Label15.Caption = ""
        frameCtaExpCC.Visible = True
        W = frameCtaExpCC.Width
        H = frameCtaExpCC.Height
        QueCombosFechaCargar "5|6|7|"
        cmbFecha(5).ListIndex = Month(vParam.fechaini) - 1
        cmbFecha(6).ListIndex = Month(vParam.fechafin) - 1
        txtAno(7).Text = Year(vParam.fechaini)
        txtAno(8).Text = Year(vParam.fechafin)
        
        
    Case 17
        Label2(26).Caption = ""
        frameCCxCta.Visible = True
        W = frameCCxCta.Width
        H = frameCCxCta.Height
        QueCombosFechaCargar "8|9|10|"
        cmbFecha(8).ListIndex = Month(vParam.fechaini) - 1
        cmbFecha(9).ListIndex = Month(vParam.fechafin) - 1
        txtAno(9).Text = Year(vParam.fechaini)
        txtAno(10).Text = Year(vParam.fechafin)
        
    Case 18, 33   '33: Legalizacion libros
    Case 19
        'DetalleExplotacion
        frameccporcta.Visible = True
        Label2(27).Visible = False
        pb7.Visible = False
        W = frameccporcta.Width
        H = frameccporcta.Height + 120
        Text3(19).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        Text3(20).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    Case 20
        'Modelo IVA 347
    Case 21
        'Cta explotacion comparativa
    Case 22, 23
        'Borre facturas cli/proveed
    Case 25, 26, 27, 39, 40
        'Balances personalizados
        chkApaisado.Value = Abs(vParam.NuevoPlanContable)
        FrameBalancesper.Visible = True
        H = FrameBalancesper.Height + 120
        W = FrameBalancesper.Width
        QueCombosFechaCargar "16|17|"
        If opcion < 39 Then
            cmbFecha(16).ListIndex = Month(vParam.fechafin) - 1
            cmbFecha(17).ListIndex = Month(vParam.fechafin) - 1
            txtAno(16).Text = Year(vParam.fechafin)
            txtAno(17).Text = Year(vParam.fechafin) - 1
            Text3(25).Text = Format(vParam.fechafin, "dd/mm/yyyy")
            If opcion > 25 Then PonerBalancePredeterminado
                
                
                
        Else
            'LEGALIZA legaliza LE-GA-LI-ZACION
                
            PonerBalancePredeterminado
            
            Text3(25).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
                
            'txtAno(0).Text = Year(CDate(RecuperaValor(Legalizacion, 2)))     'Inicio
            txtAno(16).Text = Year(CDate(RecuperaValor(Legalizacion, 3)))     'Fin
            
            'cmbFecha(0).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 2))) - 1
            cmbFecha(16).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 3))) - 1
            
            Cad = RecuperaValor(Legalizacion, 4)
            If Val(Cad) = 0 Then
                chkBalPerCompa.Value = 0
            Else
                txtAno(17).Text = Val(txtAno(16).Text) - 1
                cmbFecha(17).ListIndex = cmbFecha(16).ListIndex
                chkBalPerCompa.Value = 1
            End If
        End If
    
    End Select
    HanPulsadoSalir = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    i = opcion
    If opcion = 23 Then i = 22
    If opcion = 26 Or opcion = 27 Or opcion = 39 Or opcion = 40 Then i = 25
    If opcion = 51 Then i = 50
    If opcion = 41 Then i = 5
    If opcion = 52 Then i = 13
    If opcion = 53 Then i = 8
    If opcion = 56 Then i = 55
    
    'Legalizacion
    HanPulsadoSalir = True
    
    If opcion < 32 Or opcion > 38 Then
    
        Me.cmdCanListExtr(i).Cancel = True
        
        'Ajustaremos el boton para cancelar algunos de los listados k mas puedan costar
        AjustaBotonCancelarAccion
        cmdCancelarAccion.Visible = False
        cmdCancelarAccion.ZOrder 0
    
    End If
    Me.Width = W + 240
    Me.Height = H + 400
    
    'Añadimos ejercicios cerrados
    If EjerciciosCerrados Then Caption = Caption & "    EJERC. TRASPASADOS"
End Sub

Private Sub AjustaBotonCancelarAccion()
On Error GoTo EAj
    Me.cmdCancelarAccion.Top = cmdCanListExtr(i).Top
    Me.cmdCancelarAccion.Left = cmdCanListExtr(i).Left + 60
    cmdCancelarAccion.Width = cmdCanListExtr(i).Width
    cmdCancelarAccion.Height = cmdCanListExtr(i).Height + 30
    Exit Sub
EAj:
    MuestraError Err.Number, "Ajuste BOTON cancelar"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not HanPulsadoSalir Then Cancel = 1
    Legalizacion = ""
End Sub


Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    i = Val(Me.imgCCost(0).Tag)
    Me.txtCCost(i).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDCost(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub Image2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
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
End Sub


Private Sub imgCCost_Click(Index As Integer)
    imgCCost(0).Tag = Index
    Set frmCC = New frmCCoste
    frmCC.DatosADevolverBusqueda = "0|1|"
    frmCC.Show vbModal
    Set frmCC = Nothing
End Sub

Private Sub ImgNumBal_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
'    Set frmB = New frmBuscaGrid
''            Cad = Cad & "|"
''            Cad = Cad & mTag.Columna & "|"
''            Cad = Cad & mTag.TipoDato & "|"
''            Cad = Cad & AnchoPorcentaje & "·"
'    frmB.vCampos = "Codigo|numbalan|N|10·" & "Descripcion|nombalan|T|60·"
'    frmB.vTabla = "sbalan"
'    frmB.vSQL = ""
'    CadenaDesdeOtroForm = ""
'    '###A mano
'    frmB.vDevuelve = "0|1|"
'    frmB.vTitulo = "Balances disponibles"
'    frmB.vSelElem = 0
'    RC = Index
'    frmB.Show vbModal
'    Set frmB = Nothing
    Screen.MousePointer = vbDefault
End Sub






Private Sub List8_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub mnPrueba_Click()
    MsgBox "prueab"
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If KeyCode = 112 Then HacerF1
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
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

Private Sub txtAno_GotFocus(Index As Integer)
PonFoco txtAno(Index)
End Sub

Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtAno_KeyPress(Index As Integer, KeyAscii As Integer)
ListadoKEYpress KeyAscii
End Sub

Private Sub txtAno_LostFocus(Index As Integer)
txtAno(Index).Text = Trim(txtAno(Index).Text)
If txtAno(Index).Text = "" Then Exit Sub
If Not IsNumeric(txtAno(Index).Text) Then
    MsgBox "Campo año debe ser numérico", vbExclamation
    txtAno(Index).SetFocus
Else
    If Index = 0 Then ComprobarFechasBalanceQuitar6y7
End If
End Sub


Private Sub txtCCost_GotFocus(Index As Integer)
    PonFoco txtCCost(Index)
End Sub

Private Sub txtCCost_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtCCost_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtCCost_LostFocus(Index As Integer)
    txtCCost(Index).Text = Trim(txtCCost(Index).Text)
    If txtCCost(Index).Text = "" Then
        Me.txtDCost(Index).Text = ""
        Exit Sub
    End If
    
    SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtCCost(Index).Text, "T")
    If SQL = "" Then
        If Index > 7 Then
            MsgBox "Centro de coste NO encontrado: " & txtCCost(Index).Text, vbExclamation
            txtCCost(Index).Text = ""
            txtCCost(Index).SetFocus
        End If
    Else
        txtCCost(Index).Text = UCase(txtCCost(Index).Text)
    End If
    Me.txtDCost(Index).Text = SQL
End Sub


Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 112 Then
        HacerF1
    Else
        If KeyCode = 107 Or KeyCode = 187 Then
            KeyCode = 0
            txtCta(Index).Text = ""
            Image3_Click Index
        End If
    End If
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        Exit Sub
    End If
    
    Select Case Index
    Case 0 To 7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 23 To 30
        'NO hace falta que sean de ultimo nivel
        Cta = (txtCta(Index).Text)
                                '********
        B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
        If B = 0 Then
            MsgBox "NO existe la cuenta: " & txtCta(Index).Text, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
        Else
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
            If B = 1 Then
                DtxtCta(Index).Tag = ""
            Else
                DtxtCta(Index).Tag = SQL
            End If
            Hasta = -1
            If Index = 6 Then
                Hasta = 7
            Else
                If Index = 0 Then
                    Hasta = 1
                Else
                    If Index = 5 Then
                        Hasta = 4
                    Else
                        If Index = 23 Then Hasta = 24
                    End If
                End If
                
            End If
                
                'If txtCta(1).Text = "" Then 'ANTES solo lo hacia si el texto estaba vacio
            If Hasta >= 0 Then
                txtCta(Hasta).Text = txtCta(Index).Text
                DtxtCta(Hasta).Text = DtxtCta(Index).Text
            End If
        End If
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
    End Select
End Sub



Private Sub txtMes_GotFocus(Index As Integer)
    PonFoco txtMes(Index)
End Sub


Private Sub txtMes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtMes_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtMes_LostFocus(Index As Integer)
'Comprobar valores
    
    txtMes(Index).Text = Trim(txtMes(Index).Text)
    If txtMes(Index).Text <> "" Then
        If Not IsNumeric(txtMes(Index).Text) Then
            MsgBox "El campo no es válido: " & txtMes(Index).Text, vbExclamation
            txtMes(Index).Text = ""
            txtMes(Index).SetFocus
        End If
    End If
End Sub


Private Sub txtNumBal_GotFocus(Index As Integer)
    PonFoco txtNumBal(Index)
End Sub


Private Sub txtNumBal_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtNumBal_LostFocus(Index As Integer)
    SQL = ""
    With txtNumBal(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "Numero de balance debe de ser numérico: " & .Text, vbExclamation
                .Text = ""
            Else
                SQL = DevuelveDesdeBD("nombalan", "sbalan", "numbalan", .Text)
                If SQL = "" Then
                    MsgBox "El balance " & .Text & " NO existe", vbExclamation
                    .Text = ""
                End If
            End If
        End If
    End With
    TextDescBalance(Index).Text = SQL
End Sub

Private Sub PonerNiveles()
Dim i As Integer
Dim J As Integer


    For i = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(i)
        Cad = "Digitos: " & J
        
        'Para los de balance presupuestario
        Me.ChkCtaPre(i).Visible = True
        Me.ChkCtaPre(i).Caption = Cad
        
    Next i
    
    For i = vEmpresa.numnivel To 9
        Me.ChkCtaPre(i).Visible = False
    Next i
    
End Sub


Private Sub CargarComboFecha()
Dim J As Integer


QueCombosFechaCargar "0|1|2|"

End Sub



Private Sub GeneraSQL(Busqueda As String, vOP As Integer)
Dim SQL As String
Dim nexo As String
Dim J As Integer
Dim wildcar As String
Dim DigiTNivel As Integer
Dim IndiceAnyo As Integer
Dim IndiceMes As Integer

    SQL = ""
    nexo = ""
    For i = 1 To vEmpresa.numnivel - 1
        wildcar = ""
        
        If wildcar <> "" Then
            SQL = SQL & nexo & " (cuentas.codmacta like '" & wildcar & "')"
            nexo = " OR "
            If SQL <> "" Then SQL = "(" & SQL & ")"
        End If
    Next i


'Nexo
    Cad = "SELECT cuentas.codmacta,nommacta From "
    If vOP >= 0 Then Cad = Cad & "Conta" & vOP & "."
    Cad = Cad & "cuentas as cuentas"
    
    
        
    
    
    'MODIFICACION DE 20 OCTUBRE 2005
    Cad = Cad & ","
    If vOP >= 0 Then Cad = Cad & "Conta" & vOP & "."
    Cad = Cad & "hsaldos"
    If EjerciciosCerrados Then Cad = Cad & "1"
    Cad = Cad & " as hs WHERE "
    Cad = Cad & "cuentas.codmacta = hs.codmacta"


    Cad = Cad & " AND "
    Cad = Cad & SQL
    If Busqueda <> "" Then Cad = Cad & Busqueda
    
    
    
    'modificacion 21 Nov 2008 . MAAAAAl para años partidos
    'Rehacemos en Marzo 2009
    If opcion = 24 Then
        IndiceAnyo = 14
        IndiceMes = 14
    Else
        IndiceAnyo = 0  'Val(txtAno(0).Text)
        IndiceMes = 0 'Val(Me.cmbFecha(0).ListIndex + 1)
    End If
    
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'AÑOS NATURALES. Normal. No toco nada
        If Val(txtAno(IndiceAnyo).Text) > Year(vParam.fechaini) Then   'Pide en siguiente
            Cad = Cad & " AND (anopsald >= " & Year(vParam.fechaini) & " AND anopsald <= " & Year(vParam.fechafin) + 1 & ")"
        Else
            Cad = Cad & " AND (anopsald >= " & txtAno(IndiceAnyo).Text & " AND anopsald <= " & txtAno(IndiceAnyo + 1).Text & ")"
        End If
    
    Else
        'AÑOS PARTIDOS.
        'Si pide en ejercicio siguiente entonces hay que contemplar desde fechaini
        J = 0
        If Val(txtAno(IndiceAnyo).Text) > Year(vParam.fechafin) Then
            J = 1
        Else
            'MAYO 2009.
            
            
            'Año de fecha fin y mes mayor que fecha fin
            If Val(txtAno(IndiceAnyo).Text) = Year(vParam.fechafin) And (Me.cmbFecha(IndiceMes).ListIndex + 1) > Month(vParam.fechafin) Then J = 1
        End If
        
        
        
        'Siempre año partido
        If J = 0 Then
            
            
            
            'Buscamos mes/anyo para la fecha de inicio del balance
            If Me.cmbFecha(IndiceAnyo).ListIndex + 1 < Month(vParam.fechaini) Then
                'EL año es el anterior
                Cad = Cad & " AND ((anopsald = " & Val(txtAno(IndiceAnyo).Text) - 1 & " AND mespsald >= " & Month(vParam.fechaini) & ")"
                Cad = Cad & " OR (anopsald = " & txtAno(IndiceAnyo + 1).Text & " AND mespsald <= " & Month(vParam.fechafin) & "))"
            
            
            Else
                'Años partidos
                Cad = Cad & " AND ((anopsald = " & txtAno(IndiceAnyo).Text & " AND mespsald >= " & Month(vParam.fechaini) & ")"
                Cad = Cad & " OR (anopsald = " & txtAno(IndiceAnyo + 1).Text & " AND mespsald <= " & Month(vParam.fechafin) & "))"
            End If
        Else
            'Ha pedido de siguiente. Las cuentas las contemplo desde INICIO de ejercicio
            Cad = Cad & " AND ((anopsald = " & Year(vParam.fechaini) & " AND mespsald >= " & Month(vParam.fechaini) & ")"
            Cad = Cad & " OR (anopsald = " & txtAno(IndiceAnyo + 1).Text & " AND mespsald <= " & Me.cmbFecha(IndiceAnyo + 1).ListIndex + 1 & ")"
            'Diferencia de DOS años
            If Val(txtAno(IndiceAnyo + 1).Text) - Year(vParam.fechaini) > 1 Then Cad = Cad & " OR (anopsald = " & Year(vParam.fechaini) + 1 & ")"
            Cad = Cad & ")"
            
        End If
    End If
    
    Busqueda = Cad

End Sub



Private Sub PonerFoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function TieneCuentasEnTmpBalance(DigitosNivel As String) As Boolean
Dim Rs As ADODB.Recordset
Dim C As String

    Set Rs = New ADODB.Recordset
    TieneCuentasEnTmpBalance = False
    C = Mid("__________", 1, CInt(DigitosNivel))
    C = "Select count(*) from Usuarios.ztmpbalancesumas  where cta like '" & C & "'"
    C = C & " AND codusu = " & vUsu.Codigo
    Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If Rs.Fields(0) > 0 Then TieneCuentasEnTmpBalance = True
        End If
    End If
    Rs.Close
End Function







Public Function GeneraBalancePresupuestario() As Boolean
Dim AUx As String
Dim Importe As Currency
Dim AUX2 As String
Dim vMes  As Integer
Dim Cta As String

On Error GoTo EGeneraBalancePresupuestario
    GeneraBalancePresupuestario = False
    If Me.chkPreMensual.Value = 0 Then
        AUx = "select codmacta,sum(imppresu)  from presupuestos "
        If SQL <> "" Then AUx = AUx & " where " & SQL
        AUx = AUx & " group by codmacta"
        
        'Para el otro
        Cad = "Select SUM(impmesde),SUM(impmesha) from hsaldos where anopsald=" & i
        Cad = Cad & " and codmacta = '"
    Else
        AUx = "select codmacta,imppresu,mespresu from presupuestos where " & SQL
        If txtMes(2).Text <> "" Then AUx = AUx & " and mespresu = " & txtMes(2).Text
        AUx = AUx & " ORDER BY codmacta,mespresu"
        'para luego
        Cad = "Select impmesde,impmesha from hsaldos where anopsald=" & i
        If txtMes(2).Text <> "" Then Cad = Cad & " and mespsald = " & txtMes(2).Text
        Cad = Cad & " and codmacta = '"
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open AUx, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ningún registro a mostrar.", vbExclamation
        Rs.Close
        Exit Function
    End If
    
    'Borramos tmp de presu 2
    AUx = "DELETE FROM Usuarios.ztmppresu2 where codusu =" & vUsu.Codigo
    Conn.Execute AUx
    
    SQL = "INSERT INTO Usuarios.ztmppresu2 (codusu, codigo, cta, titulo,  mes, Presupuesto, realizado) VALUES ("
    SQL = SQL & vUsu.Codigo & ","
    
    Cont = 0
    Do
        Cont = Cont + 1
        Rs.MoveNext
    Loop Until Rs.EOF
    Rs.MoveFirst
    
    'Ponemos el PB4
    pb4.Max = Cont + 1
    pb4.Value = 0
    If Cont > 3 Then pb4.Visible = True
    Cta = ""
    Cont = 1   'Contador
    While Not Rs.EOF
        If Me.chkPreMensual.Value = 1 Then
            If Cta <> Rs!codmacta Then
                vMes = 1
                Cta = Rs!codmacta
            End If
            
            If Rs!mespresu > vMes Then
                For i = vMes To Rs!mespresu - 1
                
                    AUx = Rs!codmacta  'Aqui pondremos el nombre
                    AUx = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", AUx, "T")
                    AUx = Cont & ",'" & Rs!codmacta & "','" & DevNombreSQL(AUx) & "',"
                    AUx = AUx & i
             
                    AUx = AUx & ",0,"
                    
                    AUX2 = Cad & Rs!codmacta & "'"
                    AUX2 = AUX2 & " AND mespsald =" & i
                    
                
                
                    Importe = ImporteBalancePresupuestario(AUX2)
                    
                    AUx = AUx & TransformaComasPuntos(CStr(Importe)) & ")"
                    If Importe <> 0 Then
                        Conn.Execute SQL & AUx
                        Cont = Cont + 1
                    End If
                Next i
            End If
            
        End If
                
        
    
    
        AUx = Rs!codmacta  'Aqui pondremos el nombre
        AUx = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", AUx, "T")
        AUx = Cont & ",'" & Rs!codmacta & "','" & DevNombreSQL(AUx) & "',"
        If Me.chkPreMensual.Value = 0 Then
            AUx = AUx & "0"
        Else
            AUx = AUx & Rs!mespresu
        End If
        AUx = AUx & "," & TransformaComasPuntos(CStr(Rs.Fields(1))) & ","
        
        'SQL
        AUX2 = Cad & Rs!codmacta & "'"
        If Me.chkPreMensual.Value = 1 Then
            AUX2 = AUX2 & " AND mespsald =" & Rs!mespresu
            'AUmento el mes
            vMes = Rs!mespresu + 1
        End If
        
        
        Importe = ImporteBalancePresupuestario(AUX2)
        'Debug.Print Importe
        AUx = AUx & TransformaComasPuntos(CStr(Importe)) & ")"
        Conn.Execute SQL & AUx
        
        'Sig
        pb4.Value = pb4.Value + 1
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    
        '2013  Junio
    ' QUitaremos si asi lo pide, el saldo de la apertura
    ' Curiosamente, las 6 y 7  NO tienen apertura(perdi y ganacias)
    RC = "" 'Por si quitamos el apunte de apertura. Guardare las cuentas para buscarlas despues en la apertura
    If chkQuitarApertura.Value = 1 Then
        AUx = "SELECT cta from Usuarios.ztmppresu2 WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
        Rs.Open AUx, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            RC = RC & ", '" & Rs!Cta & "'"
            Rs.MoveNext
        Wend
        Rs.Close
        
        
        
        'Subo qui lo de quitar apertura
        If RC <> "" Then
            RC = Mid(RC, 2)
            AUx = " AND codmacta IN (" & RC & ")"
            
            Cad = "SELECT codmacta cta,sum(coalesce(timported,0))-sum(coalesce(timporteh,0)) as importe"
            Cad = Cad & " from hlinapu where codconce=970 and fechaent='" & Format(vParam.fechaini, FormatoFecha) & "'"
            Cad = Cad & AUx
            Cad = Cad & " GROUP BY 1"
            Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                Cad = "UPDATE Usuarios.ztmppresu2 SET realizado=realizado-" & TransformaComasPuntos(CStr(Rs!Importe))
                
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & Rs!Cta & "' AND mes = "
                If Me.chkPreMensual.Value = 1 Then
                    Cad = Cad & " 1"
                Else
                    Cad = Cad & " 0"
                End If
                Conn.Execute Cad
                Rs.MoveNext
            Wend
            Rs.Close
                
            
            
        End If
        
        
        
    End If
    
    
    'Si pide a 3 DIGITOS este es el momemto
    'Sera facil.
    'Hacemos un insert into con substring
 
        'SUBNIVEL
        AUx = ""
        For i = 1 To 9
            If ChkCtaPre(i).Value = 1 Then
                
                AUx = DevuelveDesdeBD("count(*)", "Usuarios.ztmppresu2", "codusu", CStr(vUsu.Codigo))
                Cont = Val(AUx)
                
                '@rownum:=@rownum+1 AS rownum      (SELECT @rownum:=0) r
                AUx = "Select " & vUsu.Codigo & " us,@rownum:=@rownum+1 AS rownum,substring(cta,1," & i & ") as cta2,mes,sum(presupuesto),sum(realizado)"
                AUx = AUx & " FROM Usuarios.ztmppresu2,(SELECT @rownum:=" & Cont & ") r WHERE codusu = " & vUsu.Codigo
                
                AUx = AUx & " AND length(cta)=" & vEmpresa.DigitosUltimoNivel
                
                AUx = AUx & " group by cta2,us,mes"
                AUx = "insert into Usuarios.ztmppresu2 (codusu, codigo, cta,   mes, Presupuesto, realizado) " & AUx
                'Insertamos
                Conn.Execute AUx
                
                'Quito los de ultimo nivel

                
                AUx = "SELECT cta from Usuarios.ztmppresu2 WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
                Rs.Open AUx, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    'Actualizo el nommacta
                    AUx = Rs!Cta  'Aqui pondremos el nombre
                    AUx = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", AUx, "T")
                    AUx = "UPDATE Usuarios.ztmppresu2  SET titulo = '" & DevNombreSQL(AUx) & "' WHERE codusu = " & vUsu.Codigo & " AND Cta = '" & Rs!Cta & "'"
                    Conn.Execute AUx
                    Rs.MoveNext
                Wend
                Rs.Close
                
                
                
            End If
        Next
        
        
        If ChkCtaPre(10).Value = 0 Then
            AUx = "DELETE FROM Usuarios.ztmppresu2 WHERE codusu = " & vUsu.Codigo & " AND cta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
            Conn.Execute AUx
        End If
        
    
    
    
  
            
        
  
    
    
    Set Rs = Nothing
    GeneraBalancePresupuestario = True
    Exit Function
EGeneraBalancePresupuestario:
    MuestraError Err.Number, "Gen. balance presupuestario"
    Set Rs = Nothing
End Function


Private Function GneraListadoPresupuesto() As Boolean

    On Error GoTo EGneraListadoPresupuesto
    GneraListadoPresupuesto = False
    If SQL <> "" Then SQL = " AND " & SQL
    SQL = "select presupuestos.* ,nommacta from presupuestos,cuentas where presupuestos.codmacta=cuentas.codmacta " & SQL
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        Rs.Close
        MsgBox "Ningun registro a listar.", vbExclamation
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztmppresu1 where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "INSERT INTO Usuarios.ztmppresu1 (codusu, codigo, cta, titulo, ano, mes, Importe) VALUES (" & vUsu.Codigo & ","
    i = i
    While Not Rs.EOF
        Cad = i & ",'" & Rs!codmacta & "','" & Rs!nommacta & "'," & Rs!anopresu
        Cad = Cad & "," & Rs!mespresu & "," & TransformaComasPuntos(CStr(Rs!imppresu)) & ")"
        Conn.Execute SQL & Cad
        'Sig
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    GneraListadoPresupuesto = True
EGneraListadoPresupuesto:
    If Err.Number <> 0 Then MuestraError Err.Number, "Listado Presupuesto"
    Set Rs = Nothing
    
End Function









'Siempre k la fecha no este en fecha siguiente
Private Function HayAsientoCierre(Mes As Byte, Anyo As Integer, Optional Contabilidad As String) As Boolean
Dim C As String
    HayAsientoCierre = False
    'C = "01/" & CStr(Me.cmbFecha(1).ListIndex + 1) & "/" & txtAno(1).Text
    C = "01/" & CStr(Mes) & "/" & Anyo
    'Si la fecha es menor k la fecha de inicio de ejercicio entonces SI k hay asiento de cierre
    If CDate(C) < vParam.fechaini Then
        HayAsientoCierre = True
    Else
        If CDate(C) > vParam.fechafin Then
            'Seguro k no hay
            Exit Function
        Else
            C = "Select count(*) from " & Contabilidad
            C = C & " hlinapu where (codconce=960 or codconce = 980) and fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            C = C & "' AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
            Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                If Not IsNull(Rs.Fields(0)) Then
                    If Rs.Fields(0) > 0 Then HayAsientoCierre = True
                End If
            End If
            Rs.Close
        End If
    End If
End Function




Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    Cad = RecuperaValor(Lista, L)
    If Cad <> "" Then
        i = Val(Cad)
        With cmbFecha(i)
            .Clear
            For Cont = 1 To 12
                RC = "25/" & Cont & "/2002"
                RC = Format(RC, "mmmm") 'Devuelve el mes
                .AddItem RC
            Next Cont
        End With
    End If
    L = L + 1
Loop Until Cad = ""
End Sub





Private Sub UltimoMesAnyoAnal1(ByRef Mes As Integer, ByRef Anyo As Integer)
    Anyo = 1900
    Mes = 13
    Set miRsAux = New ADODB.Recordset
    SQL = "select max(anoccost) from hsaldosanal1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Anyo = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    If Anyo > 1900 Then
        SQL = "select max(mesccost) from hsaldosanal1 where anoccost =" & Anyo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Mes = 1
        If Not miRsAux.EOF Then
            Mes = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
    End If
    Set miRsAux = Nothing
End Sub

Private Function GeneraCtaExplotacionCC() As Boolean
Dim RC As Byte

    GeneraCtaExplotacionCC = False
    
    
    'Borramos datos
    SQL = "Delete from Usuarios.zctaexpcc where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    

    If chkCtaExpCC(1).Value = 1 Then
        'Hacemos primero el periodo anterior
        RC = HacerCtaExploxCC(CInt(txtAno(7).Text) - 1, CInt(txtAno(8).Text) - 1)
        If RC = 0 Then Exit Function  'ha habido algun error
        
        'Si ha Generado datos los paso un momento para que el seguiente proceso no los borre
        If RC = 2 Then
            If Not ProcesoCtaExplotacionCC(0) Then Exit Function
        End If
        
    End If
    
    'Este bloque lo hace siempre
    RC = HacerCtaExploxCC(CInt(txtAno(7).Text), CInt(txtAno(8).Text))
    If RC > 0 Then
        RC = 0
        'Si era compartativo
        'Mezclamos los datos de ahora con los de antes
        If chkCtaExpCC(1).Value = 1 Then
            If Not ProcesoCtaExplotacionCC(1) Then RC = 1
        End If
        If RC = 0 Then GeneraCtaExplotacionCC = True
    End If
    

    
    'Eliminamos datos temporales
    If chkCtaExpCC(1).Value = 1 Then
        ProcesoCtaExplotacionCC 2
    End If
    

End Function

'0: Error    1: No hya datods       2: OK
Private Function HacerCtaExploxCC(Anyo1 As Integer, Anyo2 As Integer) As Byte
Dim A1 As Integer, M1 As Integer
Dim Post As Boolean

    On Error GoTo EGeneraCtaExplotacionCC
    HacerCtaExploxCC = 0
    
    UltimoMesAnyoAnal1 M1, A1
    
    'Si años consulta iguales
    If txtAno(7).Text = txtAno(8).Text Then
         Cad = " anoccost=" & Anyo1 & " AND mesccost>=" & Me.cmbFecha(5).ListIndex + 1
         Cad = Cad & " AND mesccost<=" & Me.cmbFecha(6).ListIndex + 1
         
    Else
        'Años disitintos
        'Inicio
        Cad = "( anoccost=" & Anyo1 & " AND mesccost>=" & Me.cmbFecha(5).ListIndex + 1 & ")"
        Cad = Cad & " OR ( anoccost=" & Anyo2 & " AND mesccost<=" & Me.cmbFecha(6).ListIndex + 1 & ")"
        'Por si la diferencia es mas de un año
        If Val(txtAno(8).Text) - Val(txtAno(7).Text) > 1 Then
            Cad = Cad & " OR (anoccost >" & Anyo1 & " AND anoccost < " & Anyo2 & ")"
        End If
    End If
    Cad = " (" & Cad & ")"
    
    RC = ""
    If txtCCost(2).Text <> "" Then RC = " codccost >='" & txtCCost(2).Text & "'"
    If txtCCost(3).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codccost <='" & txtCCost(3).Text & "'"
    End If
    
    
    'Si han puesto desde hasta cuenta
    If txtCta(29).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codmacta >='" & txtCta(29).Text & "'"
    End If
    
    If txtCta(30).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codmacta <='" & txtCta(30).Text & "'"
    End If
    
    
    'Cogemos prestada la tabla tmpCierre cargando las cuentas k
    'tengan en hpsaldanal y hpsaldana1 si asi lo recuieren las fechas
    SQL = "Delete  from tmpctaexpCC"
    Conn.Execute SQL
    
    
    
    'Insertamos las cuentas desde hpsald1 si hicera o hiciese falta
    Tablas = ""
    If Anyo1 < A1 Then
        Tablas = "SI"
    Else
        If Anyo1 = A1 Then
            'Dependera del mes
            If M1 > (Me.cmbFecha(5).ListIndex + 1) Then Tablas = "OK"
        End If
    End If

    If RC <> "" Then Cad = RC & " AND " & Cad
    SQL = "INSERT INTO tmpctaexpCC (codusu,cta,codccost) SELECT "
    SQL = SQL & vUsu.Codigo & ",codmacta,codccost from hsaldosanal"
    'Si es de hco
    If Tablas <> "" Then SQL = SQL & "1"
    SQL = SQL & " Where "
    SQL = SQL & Cad
    SQL = SQL & " GROUP BY codccost,codmacta"
    Conn.Execute SQL
    
    
    
    'Diciembre 2012
    'Si ha marcado "solo" centros de reparto elimino aquellos eque n
    'Borro todos aquellos cc que no sean de reparto
    If chkCtaExpCC(2).Value = 1 Then
        SQL = "DELETE FROm tmpctaexpCC WHERE codusu = " & vUsu.Codigo & " AND "
        SQL = SQL & " NOT codccost IN (select distinct(subccost) from linccost) "
        Conn.Execute SQL
    End If
    
    
    'AHora en  tenemos todas las cuentas a tratar
    'Para ello cogeremos
    SQL = "Select count(*) from tmpctaexpCC where codusu = " & vUsu.Codigo
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not Rs.EOF Then
        Cont = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close

    If Cont = 0 Then
        HacerCtaExploxCC = 1
        'MsgBox "Ningun registro a mostrar", vbExclamation
    Else
        'mes ini, ano ini, mes pedido, ano pedido, mes fin, ano fin
        'Cad = cmbFecha(5).ListIndex + 1 & "|" & txtAno(7).Text & "|"
        'Cad = Cad & cmbFecha(7).ListIndex + 1 & "|"
        Cad = cmbFecha(5).ListIndex + 1 & "|" & Anyo1 & "|"
        Cad = Cad & cmbFecha(7).ListIndex + 1 & "|"
        
        
        
        'El año del mes de calculo tiene k estar entre los años pedidos
        If cmbFecha(7).ListIndex >= cmbFecha(5).ListIndex Then
            Cad = Cad & Anyo1
        Else
            Cad = Cad & Anyo2
        End If
        Cad = Cad & "|"
        Cad = Cad & cmbFecha(6).ListIndex + 1 & "|" & Anyo2 & "|"
        
        'Ajusta los valores en modulo
        AjustaValoresCtaExpCC Cad
        
        'Si ha pediod los movimientos posteriores
        Post = (chkCtaExpCC(0).Value = 1)
        
        SQL = "Select cta,tmpctaexpCC.codccost,nommacta,nomccost from tmpctaexpCC,cuentas,cabccost where cuentas.codmacta=tmpctaexpCC.cta and cabccost.codccost=tmpctaexpCC.codccost and codusu = " & vUsu.Codigo
        'Vemos hasta donde hay de fechas en hco
        FechaFinEjercicio = CDate("01/" & M1 & "/" & A1)
        Rs.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
        While Not Rs.EOF
            
            Tablas = Rs.Fields(0) & "|" & Rs.Fields(1) & "|" & DevNombreSQL(Rs.Fields(2))
            Tablas = Tablas & "|" & DevNombreSQL(Rs.Fields(3)) & "|"
        
            'Tb ponemos la pb
            Label15.Caption = Rs.Fields(0)
            Label15.Refresh
        
            CtaExploCentroCoste Tablas, Post, FechaFinEjercicio
    
            'Siguiente
            Rs.MoveNext
        Wend
        Rs.Close
        
        
        
        
        
        A1 = 0
        SQL = "Select count(*) from Usuarios.zctaexpcc where codusu =" & vUsu.Codigo
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            A1 = DBLet(Rs.Fields(0), "N")
        End If
        Rs.Close
        If A1 = 0 Then
            'MsgBox "Ningun registro a mostrar", vbExclamation
            'GeneraCtaExplotacionCC = False
            HacerCtaExploxCC = 1
        Else
            HacerCtaExploxCC = 2
            'GeneraCtaExplotacionCC = True
        End If
        
    End If
    
    Exit Function
EGeneraCtaExplotacionCC:
    MuestraError Err.Number, "Genera Cta Explotacion CC"
End Function

'Proceso:    1.- Updatear codusu para que no borre los datos
'            2.- Mezclar los datos de otros actual /siguiente
'            3.- Borramos
Private Function ProcesoCtaExplotacionCC(Proceso As Byte) As Boolean

    ProcesoCtaExplotacionCC = False
    Select Case Proceso
    Case 0
        Cont = 0
        i = 0
        SQL = "Select min(codusu) from Usuarios.zctaexpcc"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then
                i = Rs.Fields(0)
                Cont = 1 'para indicar que no es null
            End If
        End If
        Rs.Close
        If i >= 10 Then
            i = 9
        Else
            If i = 0 Then
                'Si NO ERA NULL es que estan ocupados(cosa rara) desde el 0 al 9
                If Cont = 1 Then
                    MsgBox "Error inesperado. Descripcion: codusu entre 0..9", vbExclamation
                    Exit Function
                End If
                i = 9
            Else
                i = i - 1
            End If
        End If
        NumRegElim = i
        SQL = "UPDATE  Usuarios.zctaexpcc set codusu = " & NumRegElim & " WHERE codusu = " & vUsu.Codigo
        Conn.Execute SQL
    
    Case 1
        'UPDATEO los valores de codusu=vusu
        SQL = "UPDATE Usuarios.zctaexpcc  SET acumd=0,acumh=0,postd=0,posth=0 where codusu = " & vUsu.Codigo
        Conn.Execute SQL
        
        'Para los valores comarativos
        SQL = "UPDATE Usuarios.zctaexpcc  SET acumd=perid,acumh=perih,postd=saldod,posth=saldoh where codusu = " & NumRegElim
        Conn.Execute SQL
        SQL = "UPDATE Usuarios.zctaexpcc  SET perid=0,perih=0,saldod=0,saldoh=0 where codusu = " & NumRegElim
        Conn.Execute SQL
        
        'En RS cargo todas las referencias de codusu= vusu
        SQL = "Select * from Usuarios.zctaexpcc WHERE codusu = "
        Rs.Open SQL & vUsu.Codigo, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        'Cojere Todas las referencias de la tabla zctaexpcc para numregelim
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL & NumRegElim, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            'Busco si tiene referncia en ACTUAL
            SQL = "codmacta = '" & miRsAux!codmacta & "' AND codccost ='" & DevNombreSQL(miRsAux!codccost) & "'"
            
            
            If Not EncontrarEn_zctaexpcc(miRsAux!codmacta, UCase(miRsAux!codccost)) Then
                'Updateo solo el codusu
                Cad = "UPDATE Usuarios.zctaexpcc  SET codusu = " & vUsu.Codigo & " WHERE codusu = " & NumRegElim & " AND " & SQL
            Else
                'UPDATEO
                Cad = "UPDATE Usuarios.zctaexpcc  SET acumd=" & TransformaComasPuntos(CStr(miRsAux!acumd))
                Cad = Cad & ", acumH=" & TransformaComasPuntos(CStr(miRsAux!acumh))
                Cad = Cad & ", postd=" & TransformaComasPuntos(CStr(miRsAux!postd))
                Cad = Cad & ", posth=" & TransformaComasPuntos(CStr(miRsAux!posth))
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND " & SQL
            End If
            Conn.Execute Cad
            'Siguiente
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        Rs.Close
        
    Case 2
            'Finalmente borramos los datos de codusu=numregeleim
            SQL = "DELETE FROM Usuarios.zctaexpcc WHERE codusu = " & NumRegElim
            Conn.Execute SQL
        
            'Ahora, para que """SOLO""" aparezcan  los que tienen valor
            ' Los importes seran
            '                   MES                 SALDO
            '   actual      perid   perdih     saldod  saldoh
            '   anterior    acumd   acumh      postd   posth
        
            If optCCComparativo(1).Value Then
                'QUiero ver movimientos del periodo con lo cual me cargare aquellos
                ' que mov periodo en actual y anterior sea 0
                Cad = "acumd =0 and acumh=0 and perid=0 and perih=0"
                SQL = "DELETE FROM Usuarios.zctaexpcc WHERE codusu = " & vUsu.Codigo & " AND " & Cad
                Conn.Execute SQL
            End If
    End Select
    ProcesoCtaExplotacionCC = True
End Function

Private Function EncontrarEn_zctaexpcc(ByRef Cta As String, ByRef CC As String) As Boolean
    EncontrarEn_zctaexpcc = False
    Rs.MoveFirst
    While Not Rs.EOF
        If Rs!codmacta = Cta And Rs!codccost = CC Then
            EncontrarEn_zctaexpcc = True
            Exit Function
        Else
            Rs.MoveNext
        End If
    Wend
    
        
End Function

Private Sub InseretaDesdeHCO(ByRef Cuenta As String)
On Error Resume Next
    Conn.Execute Cuenta
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub





Private Function GeneraCCxCtaExplotacion() As Boolean
Dim A1 As Integer, M1 As Integer
Dim Post As Boolean

    On Error GoTo EGeneraCCxCtaExplotacion
    GeneraCCxCtaExplotacion = False
    
    UltimoMesAnyoAnal1 M1, A1
    
    'Si años consulta iguales
    If txtAno(9).Text = txtAno(10).Text Then
         Cad = " anoccost=" & txtAno(9).Text & " AND mesccost>=" & Me.cmbFecha(8).ListIndex + 1
         Cad = Cad & " AND mesccost<=" & Me.cmbFecha(9).ListIndex + 1
         
    Else
        'Años disitintos
        'Inicio
        Cad = "( anoccost=" & txtAno(9).Text & " AND mesccost>=" & Me.cmbFecha(8).ListIndex + 1 & ")"
        Cad = Cad & " OR ( anoccost=" & txtAno(10).Text & " AND mesccost<=" & Me.cmbFecha(9).ListIndex + 1 & ")"
        'Por si la diferencia es mas de un año
        If Val(txtAno(10).Text) - Val(txtAno(9).Text) > 1 Then
            Cad = Cad & " OR (anoccost >" & txtAno(9).Text & " AND anoccost < " & txtAno(10).Text & ")"
        End If
    End If
    Cad = " (" & Cad & ")"

    Tablas = ""
    If txtCta(14).Text <> "" Then Tablas = "codmacta >= '" & txtCta(14).Text & "'"
    If txtCta(15).Text <> "" Then
    If Tablas <> "" Then Tablas = Tablas & " AND "
     Tablas = Tablas & "codmacta <= '" & txtCta(15).Text & "'"
    End If
    
    RC = ""
    If txtCCost(4).Text <> "" Then RC = " hsaldosanal.codccost >='" & txtCCost(4).Text & "'"
    If txtCCost(5).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " hsaldosanal.codccost <='" & txtCCost(5).Text & "'"
    End If
    
    
    'Cogemos presta la tabla tmpCierre cargando las cuentas k
    'tengan en hpsaldanal y hpsaldana1 si asi lo recuieren las fechas
    SQL = "Delete  from tmpctaexpCC"
    Conn.Execute SQL
    
    If RC <> "" Then Cad = RC & " AND " & Cad
    If Tablas <> "" Then Cad = Cad & " AND " & Tablas
    
    SQL = "INSERT INTO tmpctaexpCC (codusu,cta,codccost) SELECT "
    SQL = SQL & vUsu.Codigo & ",codmacta,hsaldosanal.codccost from hsaldosanal"
    If txtAno(9).Text <= A1 Then
        If M1 <= Me.cmbFecha(9).ListIndex + 1 Then
            SQL = SQL & "1" 'ANALITICA EN CERRADOS
        End If
    End If
        
    SQL = SQL & " as hsaldosanal,cabccost where "
    SQL = SQL & " hsaldosanal.codccost = cabccost.codccost AND "
    SQL = SQL & Cad

    'Esta marcado solo los de reparto
    If Me.optCCxCta(1).Value Then SQL = SQL & " AND idsubcos <> 1"
        
    SQL = SQL & " group by codccost,codmacta"
    Conn.Execute SQL
    
    'Si estaba marcado el 2 entonces tendre k eliminar de la tabla tmpctaexpCC los datos
    'de   codccost que esten en linccost
    If Me.optCCxCta(2).Value Then
        Label2(26).Caption = "CC de reparto /"
        Label2(26).Refresh
        espera 0.2
        
        SQL = "Select distinct(subccost) from linccost"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            SQL = "DELETE FROM tmpctaexpCC where codusu = " & vUsu.Codigo & " AND codccost = '" & miRsAux.Fields(0) & "'"
            Conn.Execute SQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    'Insertamos las cuentas desde hpsald1 si hicerao hiciese falta
    Tablas = ""
    If Val(txtAno(9).Text) < A1 Then
        Tablas = "SI"
    Else
        If Val(txtAno(9).Text) = A1 Then
            'Dependera del mes
            If M1 > (Me.cmbFecha(8).ListIndex + 1) Then Tablas = "OK"
        End If
    End If
    
    
    'AHora en  tenemos todas las cuentas a tratar
    'Para ello cogeremos
    SQL = "Select count(*) from tmpctaexpCC where codusu = " & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not miRsAux.EOF Then
        Cont = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If Cont = 0 Then
        MsgBox "Ningun registro a mostrar", vbExclamation
    Else
        'mes ini, ano ini, mes pedido, ano pedido, mes fin, ano fin
        Cad = cmbFecha(8).ListIndex + 1 & "|" & txtAno(9).Text & "|"
        Cad = Cad & cmbFecha(10).ListIndex + 1 & "|"
        'El año del mes de calculo tiene k estar entre los años pedidos
        If cmbFecha(8).ListIndex >= cmbFecha(9).ListIndex Then
            Cad = Cad & txtAno(9).Text
        Else
            Cad = Cad & txtAno(10).Text
        End If
        Cad = Cad & "|"
        Cad = Cad & cmbFecha(9).ListIndex + 1 & "|" & txtAno(10).Text & "|"
        
        'Ajusta los valores en modulo
        AjustaValoresCtaExpCC Cad
        
        'Si ha pediod los movimientos posteriores
        Post = (chkCC_Cta.Value = 1)
        
        'Borramos datos
        SQL = "Delete from Usuarios.zctaexpcc where codusu = " & vUsu.Codigo
        Conn.Execute SQL
        
                
        
        SQL = "Select cta,tmpctaexpCC.codccost,nommacta,nomccost from tmpctaexpCC,cuentas,cabccost where cuentas.codmacta=tmpctaexpCC.cta and cabccost.codccost=tmpctaexpCC.codccost and codusu = " & vUsu.Codigo
        'Vemos hasta donde hay de fechas en hco
        FechaFinEjercicio = CDate("01/" & M1 & "/" & A1)
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Tablas = Rs.Fields(0) & "|" & Rs.Fields(1) & "|" & DevNombreSQL(Rs.Fields(2)) & "|" & DevNombreSQL(Rs.Fields(3)) & "|"
            
            'Tb ponemos la pb
            Label2(26).Caption = Rs.Fields(0)
            Label2(26).Refresh
    
            CtaExploCentroCoste Tablas, Post, FechaFinEjercicio
    
            'Siguiente
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        GeneraCCxCtaExplotacion = True
    End If
    
    'Contamos para ver si tiene datos
    If GeneraCCxCtaExplotacion Then
        A1 = 0
        Set miRsAux = New ADODB.Recordset
        SQL = "Select count(*) from Usuarios.zctaexpcc where codusu =" & vUsu.Codigo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            A1 = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        If A1 = 0 Then
            MsgBox "Ningun registro a mostrar", vbExclamation
            GeneraCCxCtaExplotacion = False
        End If
    End If
    Set miRsAux = Nothing
    Exit Function
EGeneraCCxCtaExplotacion:
    MuestraError Err.Number, "Genera C. coste por Cta. Explotacion" & vbCrLf & Err.Description
End Function



Private Function ObtenerDatosCCCtaExp() As Boolean

On Error GoTo EObtenerDatosCCCtaExp
    ObtenerDatosCCCtaExp = False
    
    Label2(27).Caption = "Obteniendo conjunto registros"
    Label2(27).Visible = True
    Me.Refresh

    If EjerciciosCerrados Then
        Tablas = "1"
    Else
        Tablas = ""
    End If
    Tablas = "hlinapu" & Tablas
    SQL = "Select cuentas.codmacta,cabccost.codccost,nommacta,nomccost FROM "
    SQL = SQL & Tablas
    SQL = SQL & ",cuentas,cabccost"
    SQL = SQL & " WHERE "
    SQL = SQL & Tablas & ".codmacta=cuentas.codmacta AND "
    SQL = SQL & Tablas & ".codccost=cabccost.codccost AND "
    'Fechas
    SQL = SQL & " fechaent >='" & Format(CDate(Text3(19).Text), FormatoFecha) & "'"
    SQL = SQL & " AND fechaent <='" & Format(CDate(Text3(20).Text), FormatoFecha) & "'"
    'Si ha puesto ctas
    If txtCta(16).Text <> "" Then SQL = SQL & " AND cuentas.codmacta >='" & txtCta(16).Text & "'"
    If txtCta(17).Text <> "" Then SQL = SQL & " AND cuentas.codmacta <='" & txtCta(17).Text & "'"
    'Si ha puesto CC
    If txtCCost(6).Text <> "" Then SQL = SQL & " AND " & Tablas & ".codccost >='" & txtCCost(6).Text & "'"
    If txtCCost(7).Text <> "" Then SQL = SQL & " AND " & Tablas & ".codccost <='" & txtCCost(7).Text & "'"
    
    'K codccost no sea nulo
    SQL = SQL & " AND not (" & Tablas & ".codccost is null)"
    'Agrupado
    SQL = SQL & " group by cuentas.codmacta,codccost"
    
    'Ya tenemos el SQL
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs.EOF Then
        Rs.Close
        MsgBox "Ningun dato entre estos parametros.", vbExclamation
        Exit Function
    End If

    'Preparar
    pb7.Value = 0
    pb7.Visible = True
    Label2(27).Caption = "Preparando datos"
    Me.Refresh
    
    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend

    Rs.MoveFirst
    DoEvents
    If PulsadoCancelar Then
        Rs.Close
        Exit Function
    End If
        
    'Eliminamos datos
    SQL = "DELETE FROM Usuarios.zlinccexplo Where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    SQL = "DELETE FROM Usuarios.zcabccexplo WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    FijaValoresCtapoCC FechaIncioEjercicio, CDate(Text3(19).Text), CDate(Text3(20).Text), EjerciciosCerrados
    
    
    
    DoEvents
    If PulsadoCancelar Then
        Rs.Close
        Exit Function
    End If
    
    Rs.MoveFirst
    i = 1
    While Not Rs.EOF
        DoEvents
        If PulsadoCancelar Then
                Rs.Close
            Set Rs = Nothing
            Exit Function
        End If
        'Los labels, progress y demas
        Label2(27).Caption = Rs!nommacta
        Label2(27).Refresh
        pb7.Value = CInt((i / Cont) * pb7.Max)
        'Hacer accion
        SQL = Rs!nommacta & "|" & Rs!nomccost & "|"
        Cta_por_CC Rs!codmacta, Rs!codccost, SQL
        'Siguiente
        Rs.MoveNext
        i = i + 1
    Wend
    Rs.Close
    
    
    ObtenerDatosCCCtaExp = True
    Exit Function
EObtenerDatosCCCtaExp:
    MuestraError Err.Number
End Function



Private Function UltimaFechaHcoCabapu() As Date


UltimaFechaHcoCabapu = CDate("01/12/1900")
SQL = "Select max(fechaent) from hcabapu1"
Set Rs = New ADODB.Recordset
Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields(0)) Then
        UltimaFechaHcoCabapu = Format(Rs.Fields(0), "dd/mm/yyyy")
    End If
End If
Rs.Close
Set Rs = Nothing
End Function



'Dada una fecha me da el trimestre
Private Function QueTrimestre(Fecha As Date) As Byte
Dim C As Byte
    
        C = Month(Fecha)
        If C < 4 Then
            QueTrimestre = 1
        ElseIf C < 7 Then
            QueTrimestre = 2
        ElseIf C < 10 Then
            QueTrimestre = 3
        Else
            QueTrimestre = 4
        End If
    
End Function
Private Function ExisteEntrada() As Boolean
    SQL = "Select importe from Usuarios.z347  where codusu = " & vUsu.Codigo & " and cliprov =" & Rs!cliprov & " AND nif ='" & Rs!NIF & "';"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ExisteEntrada = True
        Importe = miRsAux!Importe
    Else
        ExisteEntrada = False
    End If
    miRsAux.Close
End Function

Private Function ExisteEntradaTrimestral(ByRef I1 As Currency, ByRef I2 As Currency, ByRef i3 As Currency, ByRef i4 As Currency, ByRef I5 As Currency) As Boolean
    SQL = "Select trim1,trim2,trim3,trim4,metalico from Usuarios.z347trimestral  where codusu = " & vUsu.Codigo & " and cliprov =" & Rs!cliprov & " AND nif ='" & Rs!NIF & "';"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ExisteEntradaTrimestral = True
        I1 = miRsAux!trim1
        I2 = miRsAux!trim2
        i3 = miRsAux!trim3
        i4 = miRsAux!trim4
        I5 = DBLet(miRsAux!metalico, "N")
    Else
        ExisteEntradaTrimestral = False
        I1 = 0: I2 = 0: i3 = 0: i4 = 0: I5 = 0
    End If
    miRsAux.Close
End Function

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




Private Function ComparaFechasCombos(Indice1 As Integer, Indice2 As Integer, InCombo1 As Integer, InCombo2 As Integer) As Boolean
    ComparaFechasCombos = False
    If txtAno(Indice1).Text <> "" And txtAno(Indice2).Text <> "" Then
        If Val(txtAno(Indice1).Text) > Val(txtAno(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        Else
            If Val(txtAno(Indice1).Text) = Val(txtAno(Indice2).Text) Then
                If Me.cmbFecha(InCombo1).ListIndex > Me.cmbFecha(InCombo2).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    ComparaFechasCombos = True
End Function




Private Sub PonerBalancePredeterminado()

    'El balance de P y G tiene el campo Perdidas=1
    Select Case opcion
    Case 27, 39
        i = 1
    Case Else
        i = 0
    End Select
    If opcion >= 50 Then
        Cont = 1
    Else
        Cont = 0
    End If
    SQL = "Select * from sbalan where predeterminado = 1 AND perdidas =" & i
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Me.txtNumBal(Cont).Text = Rs.Fields(0)
        TextDescBalance(Cont).Text = Rs.Fields(1)
    End If
    Rs.Close
    Set Rs = Nothing
    Cont = 0
End Sub



Private Function ObtenerFechasEjercicioContabilidad(Inicio As Boolean, Contabi As Integer) As Date

    On Error GoTo EObtenerFechasEjercicioContabilidad
    If Inicio Then
        ObtenerFechasEjercicioContabilidad = vParam.fechaini
    Else
        ObtenerFechasEjercicioContabilidad = vParam.fechafin
    End If
    
    Rs.Open "Select fechaini,fechafin from Conta" & Contabi & ".parametros", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Inicio Then
            ObtenerFechasEjercicioContabilidad = Rs.Fields(0)
        Else
            ObtenerFechasEjercicioContabilidad = Rs.Fields(1)
        End If
        Rs.Close
    Else
        Rs.Close
        GoTo EObtenerFechasEjercicioContabilidad
    End If
    
    Exit Function
EObtenerFechasEjercicioContabilidad:
    MuestraError Err.Number, "Obtener Fechas Ejercicio Contabilidad: " & Contabi
End Function




'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'           Para la legalizacion de libros
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------

Private Sub GeneraLegalizaPRF(ByRef OtrosP As String, NumPara As Integer)
Dim NomArchivo As String

    'Estos informes los tengo k poner a mano
    'Si los cambiaramos habria k cambiarlos en imprime y aqui

    NomArchivo = App.Path & "\InformesD\"
    Select Case opcion
    Case 32
        NomArchivo = NomArchivo & "DiarioOf.rpt"
    Case 33
        NomArchivo = NomArchivo & "resumen.rpt"
    Case 34
        NomArchivo = NomArchivo & "ConsExtracL1.rpt"
    Case 35
        NomArchivo = NomArchivo & "AsientoHco.rpt"
    Case 36, 41
        NomArchivo = NomArchivo & "Sumas2.rpt"
    Case 37
        NomArchivo = NomArchivo & "faccli2.rpt"
    Case 38
        NomArchivo = NomArchivo & "facprov2.rpt"
    Case 39, 40
        If vParam.NuevoPlanContable Then
            'Nuevos balances
            If chkBalPerCompa.Value = 0 Then
                NomArchivo = NomArchivo & "balance1a.rpt"
            Else
                NomArchivo = NomArchivo & "balance2a.rpt"
            End If
        Else
            If chkBalPerCompa.Value = 0 Then
                NomArchivo = NomArchivo & "balance1.rpt"
            Else
                NomArchivo = NomArchivo & "balance2.rpt"
            End If

        End If
    End Select

   With frmVisReportN
        If opcion <> 34 Then
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
        Else
            .FormulaSeleccion = "{ado_lineas.codusu}=" & vUsu.Codigo
        End If
        .SoloImprimir = False
        .OtrosParametros = OtrosP
        .NumeroParametros = NumPara
        .MostrarTree = False
        .Informe = NomArchivo
        .ExportarPDF = True
        .Show vbModal
    End With
 
End Sub


Private Function CompararEmpresasBlancePerson(CONTA As Integer, ByRef E As Cempresa, FechaInicio As Date) As Boolean
Dim i As Integer
Dim J As Integer
    On Error GoTo ECompararEmpresasBlancePerson
    CompararEmpresasBlancePerson = False
    
    RC = "Empresa configurada"
    Set Rs = New ADODB.Recordset
    
    SQL = "Select * from Conta" & CONTA & ".Empresa"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
       Cad = "Empresa no configurada"
       Rs.Close
       Exit Function
    End If

    'OK empresa
    'Veamos los niveles
    RC = "Niveles"
    
    
    'Numero de niveles
    i = Rs!numnivel
    If i <> vEmpresa.numnivel Then
        Cad = "numero de niveles distintos. " & vEmpresa.numnivel & " - " & i
        Rs.Close
        Exit Function
    End If
    
    
    
    For J = 1 To vEmpresa.numnivel - 1
        i = DigitosNivel(J)
        NumRegElim = DBLet(Rs.Fields(3 + J), "N")
        If i <> NumRegElim Then
            Cad = "Numero de digitos de nivel " & J & " son distintos. " & i & " - " & NumRegElim
            Rs.Close
            Exit Function
        End If
    Next J
    
    Rs.Close
    
    
    RC = "Parametros"
    Set Rs = New ADODB.Recordset
    
    SQL = "Select * from Conta" & CONTA & ".Parametros"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
       Cad = "Parametros no configurados"
       Rs.Close
       Exit Function
    End If
    
    
    If Rs!fechaini <> FechaInicio Then
        Cad = "Fecha inicio ejercicios distintas: " & FechaInicio & " - " & Rs!fechaini
        Rs.Close
        Exit Function
    End If
    If FechaInicio <> vParam.fechaini Then
        Cad = "La fecha de inicio de ejercicio no coincide con los datos de la empresa actual. " & vParam.fechaini & " - " & FechaInicio
        Rs.Close
        Exit Function
    End If
    Rs.Close
    CompararEmpresasBlancePerson = True
    
ECompararEmpresasBlancePerson:
    If Err.Number <> 0 Then
        Cad = "Conta " & CONTA & vbCrLf & RC & vbCrLf & Err.Description
    Else
        Cad = ""
        CompararEmpresasBlancePerson = True
    End If
    Set Rs = Nothing
End Function





Private Sub ComprobarFechasBalanceQuitar6y7()
    On Error GoTo EComprobarFechasBalanceQuitar6y7
    If Not EjerciciosCerrados Then
    End If
    Exit Sub
EComprobarFechasBalanceQuitar6y7:
    Err.Clear
End Sub




Public Sub ListadoKEYpress(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then
        KEYpress KeyAscii
    Else
        If KeyAscii = 16 Then HacerF1
    End If
End Sub

Private Sub HacerF1()
    Select Case opcion
    Case 1
    Case 2
    
    Case 4
    Case 3
    
    Case 5
    Case 6
    Case 7
    Case 8
    Case 9
    Case 10
        cmdBalPre_Click
    Case 13
        
    Case 14
    Case 15
        cmdSaldosCC_Click
    Case 16
        cmdCtaExpCC_Click
    Case 17
        cmdCCxCta_Click
    Case 18
    Case 19
        cmdCtapoCC_Click
        
    Case 21
        
    Case 54
    Case Else
    
    End Select
End Sub






Private Function Volcar347TablaTmp2() As Boolean
Dim Imp2 As Currency
Dim CuatroImportes(3) As Currency
On Error GoTo EVolcar
    Volcar347TablaTmp2 = False


    SQL = "DELETE from Usuarios.zsimulainm where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Select * from Usuarios.z347 where codusu = " & vUsu.Codigo & " ORDER BY nif"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = "Insert Into Usuarios.zsimulainm (codusu, codigo,  nomconam,  nominmov, fechaadq, valoradq, amortacu, totalamor) VALUES (" & vUsu.Codigo & ","
    Cad = ""
    Cont = 0
    While Not Rs.EOF
        If Rs!NIF <> Cad Then
            If Cad <> "" Then
                'Es otro NIF
                'Sera insert into
                Inserta347Agencias CuatroImportes(0), CuatroImportes(2), True
                Inserta347Agencias CuatroImportes(1), CuatroImportes(3), False
            End If
            Cad = Rs!NIF
            RC = Rs!razosoci
            CuatroImportes(0) = 0: CuatroImportes(1) = 0: CuatroImportes(2) = 0: CuatroImportes(3) = 0:
        End If
        'Sera UPDATE
        Select Case Rs!cliprov
        Case 48
            i = 0
        Case 49
            i = 1
        Case 70
            i = 2
        Case 71
            i = 3
        End Select
        CuatroImportes(i) = Rs!Importe
        
        
        
        Rs.MoveNext
        
    Wend
    Rs.Close
    'Metemos el ultimo registro
    Inserta347Agencias CuatroImportes(0), CuatroImportes(2), True
    Inserta347Agencias CuatroImportes(1), CuatroImportes(3), False
    Set Rs = Nothing
    Volcar347TablaTmp2 = True
    Exit Function
EVolcar:
    MuestraError Err.Number
    Set Rs = Nothing
End Function


Private Sub Inserta347Agencias(Importe1 As Currency, importe2 As Currency, Ventas As Boolean)
Dim C As String
    'SQL = "zsimulainm
    
    If Importe1 <> 0 Or importe2 <> 0 Then
        Cont = Cont + 1
        C = Cont & ",'" & Cad & "','" & DevNombreSQL(RC) & "','"
        If Ventas Then
            C = C & "VENTAS"
        Else
            C = C & "COMPRAS"
        End If
        C = C & "'," & TransformaComasPuntos(CStr(Importe1))
        C = C & "," & TransformaComasPuntos(CStr(importe2))
        C = C & "," & TransformaComasPuntos(CStr(Importe1 + importe2)) & ")"
        C = SQL & C
       Conn.Execute C
    End If
End Sub



