VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESRemesas 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remesas"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16035
   Icon            =   "frmTESRemesas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   16035
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6930
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCreacionRemesa 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   90
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   15855
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   12060
         TabIndex        =   60
         Tag             =   "Importe|N|N|||reclama|importes|||"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Frame FrameConcepto 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         TabIndex        =   30
         Top             =   60
         Width           =   15645
         Begin VB.TextBox txtCuentas 
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
            Left            =   240
            TabIndex        =   12
            Text            =   "Text2"
            Top             =   3150
            Width           =   1335
         End
         Begin VB.TextBox txtNCuentas 
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
            Left            =   1620
            TabIndex        =   58
            Text            =   "Text2"
            Top             =   3150
            Width           =   3525
         End
         Begin VB.TextBox txtRemesa 
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
            Left            =   6630
            MaxLength       =   50
            TabIndex        =   14
            Tag             =   "Descripción|T|N|||remesas|descripción|||"
            Top             =   3150
            Width           =   5025
         End
         Begin VB.CheckBox chkAgruparRemesaPorEntidad 
            Caption         =   "Distribuir recibos por entidad"
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
            Left            =   11910
            TabIndex        =   56
            Top             =   2370
            Width           =   3315
         End
         Begin VB.CheckBox chkComensaAbonos 
            Caption         =   "Compensar abonos"
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
            Left            =   11910
            TabIndex        =   55
            Top             =   1920
            Width           =   2745
         End
         Begin VB.TextBox txtFecha 
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
            Index           =   4
            Left            =   5250
            TabIndex        =   13
            Tag             =   "Fecha Reclamación|F|N|||reclama|fecreclama|dd/mm/yyyy||"
            Text            =   "99/99/9999"
            Top             =   3150
            Width           =   1245
         End
         Begin VB.TextBox txtFecha 
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
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "imgConcepto"
            Top             =   810
            Width           =   1305
         End
         Begin VB.TextBox txtFecha 
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
            Index           =   1
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "imgConcepto"
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox txtNumFac 
            Alignment       =   1  'Right Justify
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
            Left            =   1230
            TabIndex        =   6
            Tag             =   "Nº factura|N|S|0||factcli|numfactu|0000000|S|"
            Top             =   1950
            Width           =   1275
         End
         Begin VB.TextBox txtNumFac 
            Alignment       =   1  'Right Justify
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
            Left            =   1230
            TabIndex        =   7
            Tag             =   "Nº factura|N|S|0||factcli|numfactu|0000000|S|"
            Top             =   2370
            Width           =   1275
         End
         Begin VB.TextBox txtSerie 
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
            Index           =   1
            Left            =   6210
            TabIndex        =   5
            Tag             =   "imgConcepto"
            Top             =   1200
            Width           =   765
         End
         Begin VB.TextBox txtSerie 
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
            Left            =   6210
            TabIndex        =   4
            Tag             =   "imgConcepto"
            Top             =   810
            Width           =   765
         End
         Begin VB.TextBox txtCuentas 
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
            Left            =   6150
            TabIndex        =   10
            Tag             =   "imgConcepto"
            Top             =   1950
            Width           =   1275
         End
         Begin VB.TextBox txtCuentas 
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
            Index           =   1
            Left            =   6150
            TabIndex        =   11
            Tag             =   "imgConcepto"
            Top             =   2370
            Width           =   1275
         End
         Begin VB.TextBox txtNSerie 
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
            Left            =   7020
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   810
            Width           =   4665
         End
         Begin VB.TextBox txtNSerie 
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
            Left            =   7020
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1200
            Width           =   4665
         End
         Begin VB.TextBox txtNCuentas 
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
            Left            =   7470
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1950
            Width           =   4185
         End
         Begin VB.TextBox txtNCuentas 
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
            Left            =   7470
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   2370
            Width           =   4185
         End
         Begin VB.TextBox txtFecha 
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
            Index           =   3
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "imgConcepto"
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox txtFecha 
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
            Index           =   2
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "imgConcepto"
            Top             =   810
            Width           =   1305
         End
         Begin VB.TextBox txtImporte 
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
            Index           =   1
            Left            =   3660
            TabIndex        =   9
            Tag             =   "imgConcepto"
            Top             =   2370
            Width           =   1275
         End
         Begin VB.TextBox txtImporte 
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
            Left            =   3660
            TabIndex        =   8
            Tag             =   "imgConcepto"
            Top             =   1950
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
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
            Index           =   1
            Left            =   240
            TabIndex        =   59
            Top             =   2880
            Width           =   1845
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   2
            Left            =   2220
            Top             =   2880
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Descripción"
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
            Index           =   0
            Left            =   6630
            TabIndex        =   57
            Top             =   2880
            Width           =   1245
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   4
            Left            =   6210
            Top             =   2880
            Width           =   240
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
            Left            =   5250
            TabIndex        =   54
            Top             =   2880
            Width           =   795
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
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
            Left            =   5190
            TabIndex        =   53
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
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
            Index           =   1
            Left            =   5190
            TabIndex        =   52
            Top             =   1230
            Width           =   585
         End
         Begin VB.Image imgSerie 
            Height          =   255
            Index           =   0
            Left            =   5850
            Top             =   840
            Width           =   255
         End
         Begin VB.Image imgSerie 
            Height          =   255
            Index           =   1
            Left            =   5850
            Top             =   1230
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Serie"
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
            Height          =   195
            Index           =   6
            Left            =   5160
            TabIndex        =   51
            Top             =   510
            Width           =   960
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Factura"
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
            Height          =   195
            Index           =   8
            Left            =   2700
            TabIndex        =   50
            Top             =   480
            Width           =   2280
         End
         Begin VB.Label Label3 
            Caption         =   "Nro.Factura"
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
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   49
            Top             =   1650
            Width           =   1590
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   2700
            TabIndex        =   48
            Top             =   840
            Width           =   690
         End
         Begin VB.Label lblFecha1 
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
            Index           =   4
            Left            =   2580
            TabIndex        =   47
            Top             =   3990
            Width           =   4095
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   2700
            TabIndex        =   46
            Top             =   1260
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   45
            Top             =   2010
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   44
            Top             =   2430
            Width           =   615
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   3390
            Top             =   855
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   3390
            Top             =   1260
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   255
            Index           =   0
            Left            =   5850
            Top             =   1980
            Width           =   255
         End
         Begin VB.Image imgCuentas 
            Height          =   255
            Index           =   1
            Left            =   5850
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   5160
            TabIndex        =   43
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   5160
            TabIndex        =   42
            Top             =   2010
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Cuenta Cliente"
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
            Height          =   195
            Index           =   11
            Left            =   5160
            TabIndex        =   41
            Top             =   1650
            Width           =   1890
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   960
            Top             =   1230
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   960
            Top             =   855
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   270
            TabIndex        =   40
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   270
            TabIndex        =   39
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Vencimiento"
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
            Height          =   195
            Index           =   18
            Left            =   270
            TabIndex        =   38
            Top             =   480
            Width           =   2280
         End
         Begin VB.Label Label3 
            Caption         =   "Importe Vencimiento"
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
            Height          =   315
            Index           =   14
            Left            =   2670
            TabIndex        =   37
            Top             =   1650
            Width           =   2340
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   2670
            TabIndex        =   36
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   2670
            TabIndex        =   35
            Top             =   2010
            Width           =   690
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2595
         Left            =   150
         TabIndex        =   26
         Top             =   3840
         Width           =   11655
         Begin MSComctlLib.ListView lwCobros 
            Height          =   2145
            Left            =   0
            TabIndex        =   27
            Top             =   360
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   3784
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
            NumItems        =   7
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
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Importe"
               Object.Width           =   3590
            EndProperty
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   11310
            Picture         =   "frmTESRemesas.frx":000C
            ToolTipText     =   "Puntear al Debe"
            Top             =   30
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   10950
            Picture         =   "frmTESRemesas.frx":0156
            ToolTipText     =   "Quitar al Debe"
            Top             =   30
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   180
         TabIndex        =   24
         Top             =   6360
         Width           =   1755
         Begin VB.Label lblIndicador 
            Alignment       =   2  'Center
            Caption         =   "Label2"
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
            Left            =   240
            TabIndex        =   25
            Top             =   210
            Width           =   1200
         End
      End
      Begin VB.CommandButton cmdAceptar 
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
         Index           =   0
         Left            =   13170
         TabIndex        =   15
         Top             =   6540
         Width           =   1155
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
         Left            =   14430
         TabIndex        =   16
         Top             =   6540
         Width           =   1095
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
         Left            =   12090
         TabIndex        =   61
         Top             =   3900
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6945
      Left            =   30
      TabIndex        =   18
      Top             =   30
      Visible         =   0   'False
      Width           =   15915
      Begin VB.Frame FrameBotonGnral2 
         Height          =   705
         Left            =   4020
         TabIndex        =   28
         Top             =   180
         Width           =   1095
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   210
            TabIndex        =   29
            Top             =   240
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Efectuar reclamacion "
               EndProperty
            EndProperty
         End
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
         Index           =   1
         Left            =   14730
         TabIndex        =   21
         Top             =   6270
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   19
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
         TabIndex        =   20
         Top             =   990
         Width           =   15525
         _ExtentX        =   27384
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
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   15270
         TabIndex        =   23
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
End
Attribute VB_Name = "frmTESRemesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Const SaltoLinea = """ + chr(13) + """

Private Const IdPrograma = 701


Public Tipo As Integer
Public vSQL As String
Public Opcion As Byte      ' 0.- Nueva remesa    1.- Modifcar remesa
                           ' 2.- Devolucion remesa
Public vRemesa As String   ' nºremesa|fecha remesa
Public ImporteRemesa As Currency

Public ValoresDevolucionRemesa As String
        'NOV 2009
        'antes: 4 campos     AHORA 5 campos
        'Concepto|ampliacion|
        'Concepto banco|ampliacion banco|
        'ahora+ Agrupa vtos

Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1

Dim SQL As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim Cont As Long
Dim i As Integer
Dim TotalReg As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

Dim CampoOrden As String
Dim Orden As Boolean
Dim Modo As Byte

Dim Txt33Csb As String
Dim Txt41Csb As String

Dim indice As Integer
Dim Codigo As Long

Dim SubTipo As Integer

Dim ModoInsertar As Boolean

Dim IndCodigo As Integer

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
    
        If ModoInsertar Then
            cmdAceptar(0).Caption = "&Aceptar"
            ModoInsertar = False
        End If
    
        CargaList
        
        Frame1.Visible = True
        Frame1.Enabled = True
        
        FrameCreacionRemesa.Visible = False
        FrameCreacionRemesa.Enabled = False
    Else
        Unload Me
    End If
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
    Select Case Index
        Case 0
            Select Case Modo
                Case 3  ' insertar
                    If Not DatosOK Then Exit Sub
                
                    If Not ModoInsertar Then
                        ModoInsertar = True
                        cmdAceptar(0).Caption = "C&onfirmar"
                        If SubTipo <> vbTipoPagoRemesa Then
                            'NuevaRemTalPag
                        Else
                            NuevaRem
                        End If
                    Else
                        If GenerarRemesa(0) Then
                            MsgBox "Remesa generada correctamente.", vbExclamation
                            cmdCancelar_Click (0)
                        End If
                    End If
                    
                    Screen.MousePointer = vbDefault
                    
                    
                Case 4  ' modificar
                    If DatosOK Then
                        ModificaDesdeFormulario Me
                        cmdCancelar_Click (0)
                    End If
            End Select
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean

    DatosOK = False

    If SubTipo <> vbTipoPagoRemesa Then
        'Para talones y pagares obligado la cuenta bancaria
        If txtCuentas(2).Text = "" Then
            MsgBox "Indique la cuenta bancaria", vbExclamation
            Exit Function
        End If
    End If

    'Fecha remesa tiene k tener valor
    If txtFecha(4).Text = "" Then
        MsgBox "Fecha de remesa debe tener valor", vbExclamation
        PonFoco txtFecha(4)
        Exit Function
    End If
    
    'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
    If FechaCorrecta2(CDate(txtFecha(4).Text), True) > 1 Then
        PonFoco txtFecha(4)
        Exit Function
    End If
    
    'Para talones pagares, vemos si esta configurado en parametros
    If SubTipo <> vbTipoPagoRemesa Then
'        If Me.cmbRemesa.ListIndex = 0 Then
'            SQL = "contapagarepte"
'        Else
'            SQL = "contatalonpte"
'        End If
'        SQL = DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1")
'        If SQL = "" Then SQL = "0"
'        If SQL = "0" Then
'            MsgBox "Falta configurar la opción en parametros", vbExclamation
'            Exit Sub
'        End If
    End If
    
    'mayo 2015
     If SubTipo = vbTipoPagoRemesa Then
        If vParamT.RemesasPorEntidad Then
            If chkAgruparRemesaPorEntidad.Value = 1 Then
                'Si agrupa pro entidad, necesit el banco por defacto
                If txtCuentas(2).Text = "" Then
                    MsgBox "Si agrupa por entidad debe indicar el banco por defecto", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    
    DatosOK = True

End Function

Private Sub Insertar()
Dim NumF As Long
Dim B As Boolean

    On Error GoTo eInsertar
    
    Conn.BeginTrans
    
'    NumF = SugerirCodigoSiguienteStr("reclama", "codigo")
'    Text1(5).Text = NumF
'    Codigo = Text1(5)
'    B = InsertarDesdeForm(Me)
'    If B Then InsertarLineas
    
eInsertar:
    If Err.Number = 0 And B Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
End Sub

Private Function InsertarLineas() As Boolean
Dim Rs As ADODB.Recordset
Dim CadValues As String
Dim CadInsert As String

    On Error GoTo eInsertarLineas

    InsertarLineas = False

'    CadInsert = "insert into reclama_facturas (codigo,numlinea,numserie,numfactu,fecfactu,numorden,impvenci) values "
'
'    CadValues = ""
'    For i = 1 To lwReclamCli.ListItems.Count
'        If lwReclamCli.SelectedItem.Checked Then
'            CadValues = CadValues & "(" & DBSet(Text1(5).Text, "N") & "," & DBSet(i, "N") & "," & DBSet(lwReclamCli.ListItems(i).Text, "T") & ","
'            CadValues = CadValues & DBSet(lwReclamCli.ListItems(i).SubItems(1), "N") & "," & DBSet(lwReclamCli.ListItems(i).SubItems(2), "F") & ","
'            CadValues = CadValues & DBSet(lwReclamCli.ListItems(i).SubItems(3), "N") & "," & DBSet(lwReclamCli.ListItems(i).SubItems(6), "N") & "),"
'        End If
'    Next i
'
'    If CadValues <> "" Then
'        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
'        Conn.Execute CadInsert & CadValues
'    End If
    
    InsertarLineas = True
    Exit Function
    
eInsertarLineas:
    MuestraError Err.Number, "Insertar Lineas", Err.Description
End Function

Private Sub cmdVtoDestino(Index As Integer)
    
    If Index = 0 Then
        TotalReg = 0
        If Not Me.lwCobros.SelectedItem Is Nothing Then TotalReg = Me.lwCobros.SelectedItem.Index
    
    
        For i = 1 To Me.lwCobros.ListItems.Count
            If Me.lwCobros.ListItems(i).Bold Then
                Me.lwCobros.ListItems(i).Bold = False
                Me.lwCobros.ListItems(i).ForeColor = vbBlack
                For Cont = 1 To Me.lwCobros.ColumnHeaders.Count - 1
                    Me.lwCobros.ListItems(i).ListSubItems(Cont).ForeColor = vbBlack
                    Me.lwCobros.ListItems(i).ListSubItems(Cont).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalReg > 0 Then
            i = TotalReg
            Me.lwCobros.ListItems(i).Bold = True
            Me.lwCobros.ListItems(i).ForeColor = vbRed
            For Cont = 1 To Me.lwCobros.ColumnHeaders.Count - 1
                Me.lwCobros.ListItems(i).ListSubItems(Cont).ForeColor = vbRed
                Me.lwCobros.ListItems(i).ListSubItems(Cont).Bold = True
            Next
        End If
        lwCobros.Refresh
        
        PonerFocoLw Me.lwCobros

    Else
    
'        frmTESRemesasImp.pCodigo = Me.lw1.SelectedItem
'        frmTESRemesasImp.Show vbModal

    End If
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Not Frame1.Visible Then
            If CadenaDesdeOtroForm <> "" Then
'                Text1(2).Text = CadenaDesdeOtroForm
'                Text1_LostFocus 2
            Else
'                PonFoco Text1(2)
            End If
            CadenaDesdeOtroForm = ""
        End If
        CargaList
    End If
    Screen.MousePointer = vbDefault
End Sub



    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmPpal.Icon
    
    For i = 0 To 1
        Me.imgSerie(i).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgCuentas(i).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next i
    Me.imgCuentas(2).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    
    For i = 0 To 4
        Me.imgFec(i).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next i
    
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
    
    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 47
    End With
    
    
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    
    'Limpiamos el tag
    PrimeraVez = True
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
        
    H = FrameCreacionRemesa.Height + 120
    W = FrameCreacionRemesa.Width
    
    FrameCreacionRemesa.Visible = False
    Me.Frame1.Visible = True
    
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    Me.cmdCancelar(0).Cancel = True
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    Orden = True
    
    CargaCombo
    
    If Tipo = 1 Then
        SubTipo = vbTipoPagoRemesa
    Else
        SubTipo = vbTalon
    End If
    
    
'    PonerFrameProgreso

End Sub


Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub


Private Sub Image3_Click(Index As Integer)

    Select Case Index
        Case 1 ' cuenta contable
            Screen.MousePointer = vbHourglass
            
            Set frmCtas = New frmColCtas
            RC = Index
            frmCtas.DatosADevolverBusqueda = "0|1"
            frmCtas.ConfigurarBalances = 3
            frmCtas.Show vbModal
            Set frmCtas = Nothing
    
    End Select
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim IT
Dim i As Integer
    For i = 1 To Me.lwCobros.ListItems.Count
        Set IT = lwCobros.ListItems(i)
        lwCobros.ListItems(i).Checked = (Index = 1)
        lwCobros_ItemCheck (IT)
        Set IT = Nothing
    Next i
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'FECHA FACTURA
    indice = Index
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtFecha(indice).Text <> "" Then frmF.Fecha = CDate(txtFecha(indice).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco txtFecha(indice)

End Sub

Private Sub imgCuentas_Click(Index As Integer)
    
    SQL = ""
    AbiertoOtroFormEnListado = True
    Set frmCtas = New frmColCtas
    frmCtas.DatosADevolverBusqueda = True
    frmCtas.Show vbModal
    Set frmCtas = Nothing
    If SQL <> "" Then
        Me.txtCuentas(Index).Text = RecuperaValor(SQL, 1)
        Me.txtNCuentas(Index).Text = RecuperaValor(SQL, 2)
    Else
        QuitarPulsacionMas Me.txtCuentas(Index)
    End If
    
    PonFoco Me.txtCuentas(Index)
    AbiertoOtroFormEnListado = False

End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2, 3, 4
        indice = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgSerie_Click(Index As Integer)
    IndCodigo = Index

    Set frmConta = New frmBasico
    AyudaContadores frmConta, txtSerie(Index), "tiporegi REGEXP '^[0-9]+$' = 0"
    Set frmConta = Nothing
    
    PonFoco Me.txtSerie(Index)
End Sub

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    Orden = Not Orden
    
    Select Case ColumnHeader
        Case "Código"
            CampoOrden = "remesas.codigo"
        Case "Fecha"
            CampoOrden = "remesas.fecremesa"
        Case "Cuenta"
            CampoOrden = "remesas.codmacta"
        Case "Nombre"
            CampoOrden = "cuentas.nommacta"
        Case "Año"
            CampoOrden = "remesas.anyo"
        Case "Importe"
            CampoOrden = "remesas.importe"
        Case "Descripción"
            CampoOrden = "remesas.descripcion"
        Case "Situación"
            CampoOrden = "descsituacion"
    End Select
    CargaList


End Sub

Private Sub lwCobros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim C As Currency
Dim Cobro As Boolean

    Cobro = True
    C = Item.Tag
    
    Importe = 0
    For i = 1 To lwCobros.ListItems.Count
        If lwCobros.ListItems(i).Checked Then Importe = Importe + lwCobros.ListItems(i).SubItems(6)
    Next i
    Text1(4).Text = Format(Importe, "###,###,##0.00")
    
    If ComprobarCero(Text1(4).Text) = 0 Then Text1(4).Text = ""
            
End Sub

Private Sub HacerToolBar(Boton As Integer)

    Select Case Boton
        Case 1
            BotonAnyadir
        Case 2
            BotonModificar
        Case 3
            BotonEliminar
        Case 5
'            BotonBuscar
        Case 6 ' ver todos
            CargaList
        Case 8
            'Imprimir factura
'            frmTESRemesasList.Show vbModal

    End Select
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If Me.lw1.SelectedItem = "" Then Exit Sub
        
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar la Remesa?"
    SQL = SQL & vbCrLf & "Código: " & lw1.SelectedItem.Text
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = lw1.SelectedItem.Text
        
        If ModificarCobros Then
            lw1.ListItems.Remove (lw1.SelectedItem.Index)
            If lw1.ListItems.Count > 0 Then
                lw1.SetFocus
            End If
        End If
'        CargaList
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Function ModificarCobros() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim FecUltCob As String
Dim Importe As Currency
Dim NumLinea As Integer


    ModificarCobros = False
    
    Conn.BeginTrans

    SQL = "select * from cobros where codrem = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    SQL = SQL & " and anyorem = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        If lw1.ListItems(lw1.SelectedItem.Index).SubItems(8) = "B" Then
            SQL = "select max(numlinea) from cobros_realizados where numserie = " & DBSet(Rs!NUmSerie, "T") & _
                    " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & _
                    " and numorden = " & DBSet(Rs!numorden, "N")
                    
            NumLinea = DevuelveValor(SQL)
        
            SQL = "delete from cobros_realizados where numserie = " & DBSet(Rs!NUmSerie, "T") & " and " & _
                    " numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and " & _
                    " numorden = " & DBSet(Rs!numorden, "N") & " and numlinea = " & DBSet(NumLinea, "N")
            Conn.Execute SQL
        End If
    
        SQL = "select fechaent, imporcob, numline1 from (select max(numlinea) numline1, sum(impcobro) imporcob from cobros_realizados where numserie = " & DBSet(Rs!NUmSerie, "T") & _
                " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & _
                " and numorden = " & DBSet(Rs!numorden, "N") & ") aaa, cobros_realizados " & _
                " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and numfactu = " & DBSet(Rs!NumFactu, "N") & _
                    " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and numorden = " & DBSet(Rs!numorden, "N") & _
                    " and numLinea = aaa.numline1 "
                    
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        FecUltCob = ""
        Importe = 0
        If Not Rs1.EOF Then
            FecUltCob = DBLet(Rs1!FechaEnt, "F")
            Importe = DBLet(Rs1!imporcob, "N")
            NumLinea = DBLet(Rs1!numline1, "N")
        End If
        Set Rs1 = Nothing
    
        SQL = "update cobros set fecultco = " & DBSet(FecUltCob, "F", "S")
        If Importe = 0 Then
            SQL = SQL & " , impcobro = " & ValorNulo
        Else
            SQL = SQL & " , impcobro = " & DBSet(Importe, "N", "S")
        End If
        SQL = SQL & ", tiporem = " & ValorNulo
        SQL = SQL & ", codrem = " & ValorNulo
        SQL = SQL & ", anyorem = " & ValorNulo
        SQL = SQL & ", siturem = " & ValorNulo
        SQL = SQL & " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and "
        SQL = SQL & " numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and "
        SQL = SQL & " numorden = " & DBSet(Rs!numorden, "N")
                    
        Conn.Execute SQL
    
        Rs.MoveNext
    Wend

    SQL = "delete from remesas where codigo = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    SQL = SQL & " and anyo = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Conn.Execute SQL


    Set Rs = Nothing
    ModificarCobros = True
    Conn.CommitTrans
    Exit Function
    
eModificarCobros:
    Conn.RollbackTrans
    MuestraError Err.Number, "Modificar Cobros", Err.Description
End Function
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub BotonAnyadir()

    Frame1.Visible = False
    Frame1.Enabled = False

    Me.FrameCreacionRemesa.Visible = True
    Me.FrameCreacionRemesa.Enabled = True
    
    ModoInsertar = False
    
    LimpiarCampos
'
'    Combo1.ListIndex = 0
'
    Modo = 3
    PonerModo Modo

    txtFecha(4).Text = Format(Now, "dd/mm/yyyy")

    txtCuentas(2).Text = BancoPropio
    If txtCuentas(2).Text <> "" Then
        txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "bancos", "descripcion", "codmacta", txtCuentas(2), "T")
        If txtNCuentas(2).Text = "" Then txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtCuentas(2).Text, "T")
    End If

    PonleFoco txtFecha(2)

'    Me.cmbRemesa.Clear
    If SubTipo = vbTipoPagoRemesa Then
'        cmbRemesa.AddItem "Efectos"
'        Cancelado = True
        Me.Label3(8).Caption = "Fecha factura"
        Label1(1).Caption = "Banco defecto"
        
        
        If vParamT.RemesasPorEntidad Then UltimoBancoRem True
        chkComensaAbonos.Visible = True
    Else
'        Cancelado = False
'        cmbRemesa.AddItem "Pagarés"
'        cmbRemesa.AddItem "Talones"
        Me.Label3(8).Caption = "Fecha recepcion"
        Label1(1).Caption = "Banco remesar"
    End If



End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    Me.chkComensaAbonos.Value = 0
    Me.chkAgruparRemesaPorEntidad.Value = 0
    
'    cmbRemesa.ListIndex = -1
    
    Me.lwCobros.ListItems.Clear
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub BotonModificar()

    Frame1.Visible = False
    Frame1.Enabled = False

    Me.FrameCreacionRemesa.Visible = True
    Me.FrameCreacionRemesa.Enabled = True
    
    ModoInsertar = False
    
    LimpiarCampos
'
'    Combo1.ListIndex = 0
'
    Modo = 4
    PonerModo Modo

    txtFecha(4).Text = Format(Now, "dd/mm/yyyy")

    txtCuentas(2).Text = BancoPropio
    If txtCuentas(2).Text <> "" Then
        txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "bancos", "descripcion", "codmacta", txtCuentas(2), "T")
        If txtNCuentas(2).Text = "" Then txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtCuentas(2).Text, "T")
    End If

    PonleFoco txtFecha(2)

'    Me.cmbRemesa.Clear
    If SubTipo = vbTipoPagoRemesa Then
'        cmbRemesa.AddItem "Efectos"
'        Cancelado = True
        Me.Label3(8).Caption = "Fecha factura"
        Label1(1).Caption = "Banco defecto"
        
        
        If vParamT.RemesasPorEntidad Then UltimoBancoRem True
        chkComensaAbonos.Visible = True
    Else
'        Cancelado = False
'        cmbRemesa.AddItem "Pagarés"
'        cmbRemesa.AddItem "Talones"
        Me.Label3(8).Caption = "Fecha recepcion"
        Label1(1).Caption = "Banco remesar"
    End If





End Sub




Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    PonerIndicador lblIndicador, Modo
    
'    ' la cuenta no se puede modificar pq cambiarian las líneas
'    Text1(2).Locked = (Modo = 4)
'    Text1(3).Locked = (Modo = 4)
'    Image3(1).Visible = (Modo = 3)
'    Image3(1).Enabled = (Modo = 3)
'    Me.Frame4.Enabled = (Modo = 3)
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar2 Button.Index
End Sub

Private Sub HacerToolBar2(Boton As Integer)
    Select Case Boton
        Case 1
'            frmTESRemesasEfe.Show vbModal
            CargaList
            
    End Select
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    PonFoco Text1(Index)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    If Text1(Index).Text = "" Then
        Exit Sub
    End If
    
    Select Case Index
        Case 1 ' fecha
            PonerFormatoFecha Text1(Index)
        
        Case 2 ' cuenta
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "La cuenta debe ser numérica: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                Text1(3).Text = ""
                Text1(6).Tag = Text1(6).Text
                PonFoco Text1(Index)
                
                Exit Sub
            End If
            
            Select Case Index
            Case Else
                'DE ULTIMO NIVEL
                Cta = (Text1(Index).Text)
                If CuentaCorrectaUltimoNivel(Cta, SQL) Then
                    Text1(Index).Text = Cta
                    Text1(3).Text = SQL
                Else
                    MsgBox SQL, vbExclamation
                    Text1(Index).Text = ""
                    Text1(3).Text = ""
                    Text1(Index).SetFocus
                End If
                
            End Select
        Case 4
            PonerFormatoDecimal Text1(Index), 1
    End Select
End Sub

Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If Text1(Indice1).Text <> "" And Text1(Indice2).Text <> "" Then
        L1 = Len(Text1(Indice1).Text)
        L2 = Len(Text1(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(Text1(Indice1).Text & "000000000", 1, L1)) > Val(Mid(Text1(Indice2).Text & "0000000000", 1, L1)) Then
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
'    Me.lblPPAL.Caption = TEXTO
'    Me.lbl2.Caption = ""
'    Me.ProgressBar1.Value = 0
'    Me.FrameProgreso.Visible = True
    Me.Refresh
End Sub


Private Sub PonerVtosRemesa(vSQL As String, Modificar As Boolean)
Dim IT
Dim ImporteTot As Currency

    lwCobros.ListItems.Clear
    If Not Modificar Then Text1(4).Text = ""
    
    ImporteTot = 0
    
    Set Me.lwCobros.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select cobros.*,nomforpa " & vSQL
        
'    If SubTipo = vbTipoPagoRemesa Then
'        SQL = " formapago.tipforpa = " & vbTipoPagoRemesa
'    Else
''--
''        If Me.cmbRemesa.ListIndex = 0 Then
''            SQL = " talon = 0"
''        Else
''            SQL = " talon = 1"
''        End If
'
'    End If
'
'    If SubTipo = vbTipoPagoRemesa Then
'        'Del efecto
'        If txtFecha(2).Text <> "" Then SQL = SQL & " AND cobros.fecvenci >= '" & Format(txtFecha(2).Text, FormatoFecha) & "'"
'        If txtFecha(3).Text <> "" Then SQL = SQL & " AND cobros.fecvenci <= '" & Format(txtFecha(3).Text, FormatoFecha) & "'"
'    Else
'        'de la recepcion de factura
'        If txtFecha(2).Text <> "" Then SQL = SQL & " AND fechavto >= '" & Format(txtFecha(2).Text, FormatoFecha) & "'"
'        If txtFecha(3).Text <> "" Then SQL = SQL & " AND fechavto <= '" & Format(txtFecha(3).Text, FormatoFecha) & "'"
'    End If
'
'    'Si ha puesto importe desde Hasta
'    If txtImporte(0).Text <> "" Then SQL = SQL & " AND impvenci >= " & TransformaComasPuntos(ImporteFormateado(txtImporte(0).Text))
'    If txtImporte(1).Text <> "" Then SQL = SQL & " AND impvenci <= " & TransformaComasPuntos(ImporteFormateado(txtImporte(1).Text))
'
'
'    'Desde hasta cuenta
'    If SubTipo = vbTipoPagoRemesa Then
'        If Me.txtCuentas(0).Text <> "" Then SQL = SQL & " AND cobros.codmacta >= '" & txtCuentas(0).Text & "'"
'        If Me.txtCuentas(1).Text <> "" Then SQL = SQL & " AND cobros.codmacta <= '" & txtCuentas(1).Text & "'"
'        'El importe
'        SQL = SQL & " AND impvenci > 0"
'
'        'MODIFICACION DE 2 DICIEMBRE del 05
'        '------------------------------------
'        'Hay un campo que indicara si el vto se remesa o NO
'        SQL = SQL & " AND noremesar=0"
'
'
'        'Si esta en situacion juridica TAMPOCO se remesa
'        SQL = SQL & " AND situacionjuri=0"
'
'        'JUNIO 2010
'        'Si tiene algio  cobrado NO dejo remesar
'        SQL = SQL & " AND impcobro is null"
'    End If
'
'
'    'Marzo 2015
'    'Comprobar
'
'
'    'Modificacion 28 Abril 06
'    '------------------------
'    ' Es para acotar mas el conjunto de recibos a remesar
'    'Serie
'    If SubTipo = vbTipoPagoRemesa Then
'        If txtSerie(0).Text <> "" Then _
'            SQL = SQL & " AND cobros.numserie >= '" & txtSerie(0).Text & "'"
'        If txtSerie(1).Text <> "" Then _
'            SQL = SQL & " AND cobros.numserie <= '" & txtSerie(1).Text & "'"
'
'        'Fecha factura
'        If txtFecha(0).Text <> "" Then _
'            SQL = SQL & " AND cobros.fecfactu >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
'        If txtFecha(1).Text <> "" Then _
'            SQL = SQL & " AND cobros.fecfactu <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
'
'        'Codigo factura
'        If txtNumFac(0).Text <> "" Then _
'            SQL = SQL & " AND scobro.numfactu >= '" & txtNumFac(0).Text & "'"
'        If txtNumFac(1).Text <> "" Then _
'            SQL = SQL & " AND scobro.numfactu <= '" & txtNumFac(1).Text & "'"
'    Else
'        'Fecha factura
'        If txtFecha(0).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
'        If txtFecha(1).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
'    End If
'
'    If SQL <> "" Then Cad = Cad & " and " & SQL
        
    Cad = Cad & " ORDER BY fecvenci"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!NumFactu, "0000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.SubItems(4) = miRsAux!FecVenci
        IT.SubItems(5) = miRsAux!nomforpa
    
        IT.Checked = True
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        IT.SubItems(6) = Format(Importe, FormatoImporte)
        
        ImporteTot = ImporteTot + Importe

        IT.Tag = Abs(Importe)  'siempre valor absoluto
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Text1(4).Text = Format(ImporteTot, "###,###,##0.00")
    

End Sub


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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
    
        Importe = DBLet(Rs!ImpVenci) + DBLet(Rs!Gastos, "N") - DBLet(Rs!impcobro, "N")
        
        SQL2 = "select max(numlinea) from cobros_realizados where numserie = " & DBSet(Rs!NUmSerie, "T") & " and numfactu = " & DBSet(Rs!NumFactu, "N")
        SQL2 = SQL2 & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and numorden = " & DBSet(Rs!numorden, "N")
        NumLin = DevuelveValor(SQL2)
        NumLin = NumLin + 1
    
        CadValues = CadValues & "(" & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N")
        CadValues = CadValues & "," & DBSet(NumLin, "N") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Now, "FH") & "," & DBSet(Importe, "N") & ",0),"
        
        
        ' actualizamos la cabecera del cobro pq ya no lo eliminamos
        SQL = "update cobros set situacion = 2, impcobro = impvenci + coalesce(gastos,0) where numserie = " & DBSet(Rs!NUmSerie, "T")
        SQL = SQL & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and numorden = " & DBSet(Rs!numorden, "N")
        
        Conn.Execute SQL
        
        Rs.MoveNext
    Wend
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute CadInsert & CadValues
    End If
    
    
    Set Rs = Nothing
    Exit Function
    
eInsertarCobrosRealizados:
    InsertarCobrosRealizados = False
End Function




Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    SQL = ""
    For i = 1 To Me.lwCobros.ListItems.Count
        If Me.lwCobros.ListItems(i).Checked Then
        
            Insertar = True
            If Me.lwCobros.ListItems(i).Bold Then
                RegistroDestino = i
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                SQL = SQL & ", ('" & lwCobros.ListItems(i).Text & "'," & lwCobros.ListItems(i).SubItems(1)
                SQL = SQL & ",'" & Format(lwCobros.ListItems(i).SubItems(2), FormatoFecha) & "'," & lwCobros.ListItems(i).SubItems(3) & ")"
            End If
            
        End If
    Next
    SQL = Mid(SQL, 2)
            
End Sub


Private Sub FijaCadenaSQLCobrosCompen()

    Cad = "numserie, numfactu, fecfactu, numorden "
    
'    cad = "numserie , numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, ctabanc1,"
'    cad = cad & "entidad, oficina, control, cuentaba, iban, fecultco, impcobro, emitdocum, "
'    cad = cad & "recedocu, contdocu, text33csb, text41csb, "
'    cad = cad & "ultimareclamacion, agente, departamento, tiporem, CodRem, AnyoRem,"
'    cad = cad & "siturem, Gastos, Devuelto, situacionjuri, noremesar, observa, transfer, referencia,"
'    cad = cad & "nomclien, domclien, pobclien, cpclien, proclien, referencia1, referencia2,"
'    cad = cad & "feccomunica, fecprorroga, fecsiniestro, fecejecutiva, nifclien, codpais, situacion  "
    
End Sub



Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And Modo = 2
    
        Toolbar2.Buttons(1).Enabled = DBLet(Rs!especial, "N")
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Sub CargaList()
Dim IT

    lw1.ListItems.Clear
    Set Me.lw1.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select wtiporemesa2.DescripcionT,remesas.codigo,remesas.anyo, remesas.fecremesa, wtiporemesa.descripcion aaa,descsituacion,remesas.codmacta,cuentas.nommacta,"
    Cad = Cad & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    Cad = Cad & " from cuentas,usuarios.wtiporemesa2,usuarios.wtiposituacionrem,remesas left join usuarios.wtiporemesa on remesas.tipo=wtiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    Cad = Cad & " and situacio=situacion and wtiporemesa2.tipo=remesas.tiporem"
    
    Cad = Cad & PonerOrdenFiltro
    
    If CampoOrden = "" Then CampoOrden = "remesas.codigo"
    Cad = Cad & " ORDER BY remesas.anyo desc, " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
    lw1.ColumnHeaders.Clear
    
    lw1.ColumnHeaders.Add , , "Código", 950
    lw1.ColumnHeaders.Add , , "Año", 700
    lw1.ColumnHeaders.Add , , "Fecha", 1350
    lw1.ColumnHeaders.Add , , "Situación", 2540
    lw1.ColumnHeaders.Add , , "Cuenta", 1440
    lw1.ColumnHeaders.Add , , "Nombre", 2940
    lw1.ColumnHeaders.Add , , "Descripción", 3340
    lw1.ColumnHeaders.Add , , "Importe", 1940, 1
    lw1.ColumnHeaders.Add , , "T", 0, 1
    
    
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = DBLet(miRsAux!Codigo, "N")
        IT.SubItems(1) = DBLet(miRsAux!Anyo, "N")
        IT.SubItems(2) = Format(miRsAux!fecremesa, "dd/mm/yyyy")
        IT.SubItems(3) = DBLet(miRsAux!descsituacion, "T")
        IT.ListSubItems(3).ToolTipText = DBLet(miRsAux!descsituacion, "T")
        IT.SubItems(4) = miRsAux!codmacta
        IT.SubItems(5) = DBLet(miRsAux!Nommacta, "T")
        IT.ListSubItems(5).ToolTipText = DBLet(miRsAux!Nommacta, "T")
        IT.SubItems(6) = DBLet(miRsAux!Descripcion, "T")
        IT.ListSubItems(6).ToolTipText = DBLet(miRsAux!Descripcion, "T")
        IT.SubItems(7) = Format(miRsAux!Importe, "###,###,##0.00")
        IT.SubItems(8) = miRsAux!Situacion
        
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



Private Function PonerOrdenFiltro()
Dim C As String
    'Filtro
    If Tipo = 1 Then
        'REMESAS
        C = RemesaSeleccionTipoRemesa(True, False, False)
    Else
'        'TALON Y PAGARE
'        If Not Me.mnFiltro1(2).Checked And Not Me.mnFiltro1(3).Checked Then
'             Me.mnFiltro1(2).Checked = True
'              Me.mnFiltro1(3).Checked = True
'        End If
'        C = RemesaSeleccionTipoRemesa(False, Me.mnFiltro1(2).Checked, Me.mnFiltro1(3).Checked)
    End If
    
    If C <> "" Then C = " AND " & C
'    Select Case Ordenacion
'    Case 1
'        PonerOrdenFiltro = "tiporem,anyo asc , codigo asc"
'        'Tipo, codigo, año (Asc)   0 y 1desc
'    Case 3
'        PonerOrdenFiltro = "anyo desc , codigo desc,tiporem"
'    Case 4
'        PonerOrdenFiltro = "anyo asc , codigo asc,tiporem"
'
'    Case Else
'        'Por defecto
'        PonerOrdenFiltro = "tiporem,anyo desc , codigo desc"
'    End Select
'    PonerOrdenFiltro = C & " ORDER BY " & PonerOrdenFiltro
    PonerOrdenFiltro = C
End Function


Private Sub CargaCombo()
'    Combo1.Clear
'    Combo1.AddItem "Carta"
'    Combo1.ItemData(Combo1.NewIndex) = 0
'    Combo1.AddItem "Email"
'    Combo1.ItemData(Combo1.NewIndex) = 1
'    Combo1.AddItem "Teléfono"
'    Combo1.ItemData(Combo1.NewIndex) = 2
    
End Sub


Private Sub NuevaRem()

Dim Forpa As String
Dim Cad As String
Dim Impor As Currency
Dim colCtas As Collection

    If SubTipo = vbTipoPagoRemesa Then
        SQL = " formapago.tipforpa = " & vbTipoPagoRemesa
    Else
'--
'        If Me.cmbRemesa.ListIndex = 0 Then
'            SQL = " talon = 0"
'        Else
'            SQL = " talon = 1"
'        End If
    
    End If
    
    If SubTipo = vbTipoPagoRemesa Then
        'Del efecto
        If txtFecha(2).Text <> "" Then SQL = SQL & " AND cobros.fecvenci >= '" & Format(txtFecha(2).Text, FormatoFecha) & "'"
        If txtFecha(3).Text <> "" Then SQL = SQL & " AND cobros.fecvenci <= '" & Format(txtFecha(3).Text, FormatoFecha) & "'"
    Else
        'de la recepcion de factura
        If txtFecha(2).Text <> "" Then SQL = SQL & " AND fechavto >= '" & Format(txtFecha(2).Text, FormatoFecha) & "'"
        If txtFecha(3).Text <> "" Then SQL = SQL & " AND fechavto <= '" & Format(txtFecha(3).Text, FormatoFecha) & "'"
    End If
    
    'Si ha puesto importe desde Hasta
    If txtImporte(0).Text <> "" Then SQL = SQL & " AND impvenci >= " & TransformaComasPuntos(ImporteFormateado(txtImporte(0).Text))
    If txtImporte(1).Text <> "" Then SQL = SQL & " AND impvenci <= " & TransformaComasPuntos(ImporteFormateado(txtImporte(1).Text))
    
    
    'Desde hasta cuenta
    If SubTipo = vbTipoPagoRemesa Then
        If Me.txtCuentas(0).Text <> "" Then SQL = SQL & " AND cobros.codmacta >= '" & txtCuentas(0).Text & "'"
        If Me.txtCuentas(1).Text <> "" Then SQL = SQL & " AND cobros.codmacta <= '" & txtCuentas(1).Text & "'"
        'El importe
        SQL = SQL & " AND (impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) > 0"
        
        
        
        'MODIFICACION DE 2 DICIEMBRE del 05
        '------------------------------------
        'Hay un campo que indicara si el vto se remesa o NO
        SQL = SQL & " AND noremesar=0"


        'Si esta en situacion juridica TAMPOCO se remesa
        SQL = SQL & " AND situacionjuri=0"

'dejo remesar lo pendiente
'        'JUNIO 2010
'        'Si tiene algio  cobrado NO dejo remesar
'        SQL = SQL & " AND impcobro is null"
    End If
    

    'Marzo 2015
    'Comprobar
    
    
    'Modificacion 28 Abril 06
    '------------------------
    ' Es para acotar mas el conjunto de recibos a remesar
    'Serie
    If SubTipo = vbTipoPagoRemesa Then
        If txtSerie(0).Text <> "" Then _
            SQL = SQL & " AND cobros.numserie >= '" & txtSerie(0).Text & "'"
        If txtSerie(1).Text <> "" Then _
            SQL = SQL & " AND cobros.numserie <= '" & txtSerie(1).Text & "'"
        
        'Fecha factura
        If txtFecha(0).Text <> "" Then _
            SQL = SQL & " AND cobros.fecfactu >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then _
            SQL = SQL & " AND cobros.fecfactu <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        
        'Codigo factura
        If txtNumFac(0).Text <> "" Then _
            SQL = SQL & " AND cobros.numfactu >= '" & txtNumFac(0).Text & "'"
        If txtNumFac(1).Text <> "" Then _
            SQL = SQL & " AND cobros.numfactu <= '" & txtNumFac(1).Text & "'"
    
    Else
        'Fecha factura
        If txtFecha(0).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    
    End If
    
    SQL = SQL & " and situacion in (0,1) "
     
    Screen.MousePointer = vbHourglass
    Set Rs = New ADODB.Recordset
    
    'Marzo 2015
    'Ver si entre los desde hastas hay importes negativos... ABONOS
    
    If SubTipo = vbTipoPagoRemesa Then
    
        'Vemos las cuentas que vamos a girar . Sacaremos codmacta
        Cad = SQL
        Cad = "cobros.codmacta=cuentas.codmacta AND (siturem is null) AND " & Cad
        Cad = Cad & " AND cobros.codforpa = formapago.codforpa ORDER BY codmacta,numfactu "
        Cad = "Select distinct cobros.codmacta FROM cobros,cuentas,formapago WHERE " & Cad
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set colCtas = New Collection
        While Not Rs.EOF
            colCtas.Add CStr(Rs!codmacta)
            Rs.MoveNext
        Wend
        Rs.Close
        
        'Ahora veremos los negativos, de las cuentas que vamos a girar
        'Sol el select de los negativos , sin numserie ni na de na
        Cad = "(impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) < 0"
        Cad = "cobros.codmacta=cuentas.codmacta AND (siturem is null) AND " & Cad
        Cad = Cad & " AND cobros.codforpa = formapago.codforpa  "
        Cad = Cad & " AND cobros.situacion in (0,1) " '++
        Cad = "Select cobros.codmacta,nommacta,numserie,numfactu,impvenci FROM cobros,cuentas,formapago WHERE " & Cad
        
        
        If colCtas.Count > 0 Then
            Cad = Cad & " AND cobros.codmacta IN ("
            For i = 1 To colCtas.Count
                If i > 1 Then Cad = Cad & ","
                Cad = Cad & "'" & colCtas.Item(i) & "'"
            Next
            Cad = Cad & ") ORDER BY codmacta,numfactu"
        
            'Seguimos
        
            Set colCtas = Nothing
            Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Cad = ""
            i = 0
            Set colCtas = New Collection
            While Not Rs.EOF
                If i < 15 Then
                    Cad = Cad & vbCrLf & Rs!codmacta & " " & Rs!Nommacta & "  " & Rs!NUmSerie & Format(Rs!NumFactu, "000000") & "   -> " & Format(Rs!ImpVenci, FormatoImporte)
                End If
                i = i + 1
                colCtas.Add CStr(Rs!codmacta)
                Rs.MoveNext
            Wend
            Rs.Close
            
            If Cad <> "" Then
                If Me.chkComensaAbonos.Value = 0 Then
                
                    If i >= 15 Then Cad = Cad & vbCrLf & "....  y " & i & " vencimientos más"
                    Cad = "Clientes con abonos. " & vbCrLf & Cad & " ¿Continuar?"
                    If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then
                        Set Rs = Nothing
                        Set colCtas = Nothing
                        Exit Sub
                    End If
                            
                Else
                    '-------------------------------------------------------------------------
                    CadenaDesdeOtroForm = ""
                    For i = 1 To colCtas.Count
                    
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "'" & colCtas.Item(i) & "',"
'--
'                        frmListado.Opcion = 36
'                        frmListado.Show vbModal
                    Next
                    frmTESCompensaAboCli.Show vbModal
                    
                    
                    CadenaDesdeOtroForm = ""
                    
                    'Actualice BD
                    Screen.MousePointer = vbHourglass
                    espera 1
                    Screen.MousePointer = vbHourglass
                    Conn.Execute "commit"
                    espera 1
                    
                End If
            End If 'colcount
        End If
        Set colCtas = Nothing
    End If
        
    
    'Que la cuenta NO este bloqueada
    i = 0
    If SubTipo = vbTipoPagoRemesa Then
        Cad = " FROM cobros,formapago,cuentas WHERE cobros.codforpa = formapago.codforpa AND (siturem is null) AND situacion = 0 and "
        Cad = Cad & " cobros.codmacta=cuentas.codmacta AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(txtFecha(4).Text), FormatoFecha) & "') AND "
        Cad = "Select cobros.codmacta,nommacta,fecbloq" & Cad & SQL & " GROUP BY 1 ORDER BY 1"
        
    Else
'        Cad = "select cuentas.codmacta,nommacta from "
'        Cad = Cad & "scarecepdoc,cuentas where scarecepdoc.codmacta=cuentas.codmacta"
'        Cad = Cad & " AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(txtFecha(4).Text), FormatoFecha) & "') "
'        Cad = Cad & " AND " & SQL & " GROUP by 1"
    End If
    
    
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Cad = ""
        i = 1
        While Not Rs.EOF
            Cad = Cad & Rs!codmacta & " - " & Rs!Nommacta & " : " & Rs!FecBloq & vbCrLf
            Rs.MoveNext
        Wend
    End If

    Rs.Close
    
    If i > 0 Then
        Cad = "Las siguientes cuentas estan bloqueadas." & vbCrLf & String(60, "-") & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    
    If SubTipo = vbTipoPagoRemesa Then
        'Efectos bancario
    
        Cad = " FROM cobros,formapago,cuentas WHERE cobros.codforpa = formapago.codforpa AND (siturem is null) AND "
        Cad = Cad & " cobros.codmacta=cuentas.codmacta AND situacion in (0,1) and "
    Else
'--
'        'Talon / Pagare
'        Cad = " FROM scarecepdoc,cuentas where scarecepdoc.codmacta=cuentas.codmacta AND"
    End If
    'Hacemos un conteo
    Rs.Open "SELECT Count(*) " & Cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        i = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Cad = Cad & SQL
    
    
    
    If i > 0 Then
        If SubTipo <> vbTipoPagoRemesa Then
'--
'            'Para talones y pagares comprobaremos que
'            'si esta configurado para contabilizar contra cta puente
'            'entonces tiene la marca
'            'PAGARE. Ver si tiene cta puente pagare
'            If Me.cmbRemesa.ListIndex = 0 Then
'                If Not vParam.PagaresCtaPuente Then i = 0
'            Else
'                If Not vParam.TalonesCtaPuente Then i = 0
'            End If
'            If i = 0 Then
'                'NO contabilizaq contra cuenta puente
'
'            Else
'                'Comrpobaremos que todos los vtos estan en contabilizados.
'                'Por eso la marca
'
'                SQL = "(select numserie,codfaccl,fecfaccl,numorden " & Cad & ")"
'                SQL = "select distinct(id) from slirecepdoc where (numserie,numfaccl,fecfaccl,numvenci) in " & SQL
'                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                SQL = ""
'                While Not RS.EOF
'                    SQL = SQL & ", " & RS!Id
'                    RS.MoveNext
'                Wend
'                RS.Close
'                'Ya tengo el numero de las recepciones
'                If SQL = "" Then
'                    'ummmmmmmm, n deberia haber pasado
'
'                Else
'                    SQL = "(" & Mid(SQL, 3) & ")"
'                    SQL = "SELECT * from scarecepdoc where Contabilizada=0 and codigo in " & SQL
'                    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'                    SQL = ""
'                    While Not RS.EOF
'                        SQL = SQL & vbCrLf & Format(RS!Codigo, "0000") & "         " & RS!codmacta & "    " & Format(RS!fecharec, "dd/mm/yyyy") & "   " & RS!numeroref
'                        RS.MoveNext
'                    Wend
'                    RS.Close
'                    If SQL <> "" Then
'                        'Hay taloes / pagares que estan recepcionados y o estan contabilizados
'                        SQL = String(70, "-") & SQL
'                        SQL = vbCrLf & "Codigo      Cuenta            Fecha         Referencia " & vbCrLf & SQL
'                        SQL = "Hay talones / pagares que estan recepcionados pero no estan contabilizados" & vbCrLf & vbCrLf & SQL
'                        MsgBox SQL, vbExclamation
'                        Set RS = Nothing
'                        Screen.MousePointer = vbDefault
'                        Exit Sub
'                    End If
'                End If
'
'            End If
        End If
        i = 1  'Para que siga por abajo
        
    End If
    
    

    'La suma
    If i > 0 Then
        SQL = "select sum(impvenci),sum(impcobro),sum(gastos) " & Cad
        Impor = 0
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then Impor = DBLet(Rs.Fields(0), "N") - DBLet(Rs.Fields(1), "N") + DBLet(Rs.Fields(2), "N")
        Rs.Close
        If Impor = 0 Then i = 0
    End If
        

    Set Rs = Nothing
    
    If i = 0 Then
        MsgBox "Ningun dato a remesar con esos valores", vbExclamation
    Else
         
         
        'Preparamos algunas cosillas
        'Aqui guardaremos cuanto llevamos a cada banco
        SQL = "Delete from tmpCierre1 where codusu =" & vUsu.Codigo
        Conn.Execute SQL
        
        'Si son talones o pagares NO hay reajuste en bancos
        'Con lo cual cargare la tabla con el banco
        
        If SubTipo <> vbTipoPagoRemesa Then
            ' Metermos cta banco, nºremesa . El resto no necesito
            SQL = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES ("
            SQL = SQL & vUsu.Codigo & ",'" & txtCuentas(2).Text & "','"
            'ANTES
            'SQL = SQL & DevNombreSQL(Me.txtDescCta(3).Text) & "'," & TransformaComasPuntos(CStr(Impor)) & ")"
            'AHora.
            SQL = SQL & txtRemesa.Text & "',0)"
            Conn.Execute SQL
        Else
            If Not chkAgruparRemesaPorEntidad.Visible Then Me.chkAgruparRemesaPorEntidad.Value = 0
            SQL = Cad 'Le paso el SELECT
'--
            If Me.chkAgruparRemesaPorEntidad.Value = 1 Then SQL = SQL & VencimientosPorEntidadBancaria
                                
        End If
        
        
        If SubTipo = vbTipoPagoRemesa Then
            PonerVtosRemesa SQL, True
        Else
        
        End If
        
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub UltimoBancoRem(Leer As Boolean)
Dim NF As Integer
Dim Cad As String
Dim Cad2 As String

On Error GoTo EUltimoBancoRem


    Cad = App.Path & "\control.dat"
    
    If Leer Then
        If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            CadenaControl = Cad
            'El primer pipe es el usuario
            txtCuentas(2).Text = RecuperaValor(Cad, 7)
        End If
    Else 'Escribir
        NF = FreeFile
        Open Cad For Output As #NF
        Cad2 = txtCuentas(2)
        Print #NF, InsertaValor(CadenaControl, 7, Cad2)
        Close #NF
    End If
EUltimoBancoRem:
    Err.Clear
End Sub

Private Sub LanzaFormAyuda(Nombre As String, indice As Integer)
    Select Case Nombre
    Case "imgSerie"
        imgSerie_Click indice
    Case "imgFecha"
        imgFec_Click indice
    Case "imgCuentas"
        imgCuentas_Click indice
    End Select
End Sub


Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub

Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtCuentas(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub txtCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCuentas_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    
    
    If txtCuentas(Index).Text = "" Then
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCuentas(Index).Text) Then
        If InStr(1, txtCuentas(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCuentas(Index).Text, vbExclamation
        txtCuentas(Index).Text = ""
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 2 'cuentas
            Cta = (txtCuentas(Index).Text)
                                    '********
            B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
            If B = 0 Then
                MsgBox "NO existe la cuenta: " & txtCuentas(Index).Text, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
            Else
                txtCuentas(Index).Text = Cta
                txtNCuentas(Index).Text = SQL
                If B = 1 Then
                    txtNCuentas(Index).Tag = ""
                Else
                    txtNCuentas(Index).Tag = SQL
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
                    txtCuentas(Hasta).Text = txtCuentas(Index).Text
                    txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
                End If
            End If
    
    End Select
    
End Sub



Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtFecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoFecha txtFecha(Index)
End Sub

Private Sub txtSerie_GotFocus(Index As Integer)
    ConseguirFoco txtSerie(Index), 3
End Sub

Private Sub txtSerie_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtSerie_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtSerie(Index).Text = UCase(Trim(txtSerie(Index).Text))
    
    If txtSerie(Index).Text = "" Then
        txtNSerie(Index).Text = ""
        Exit Sub
    End If
    
    Select Case Index
        Case 0, 1 'tipos de movimiento
            txtNSerie(Index).Text = DevuelveDesdeBD("nomregis", "contadores", "tiporegi", txtSerie(Index), "T")
    End Select
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

End Sub


Private Sub txtNumFac_GotFocus(Index As Integer)
    ConseguirFoco txtNumFac(Index), 3
End Sub

Private Sub txtNumFac_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtNumFac_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtNumFac_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtNumFac(Index).Text = UCase(Trim(txtNumFac(Index).Text))
    
    
    Select Case Index
        Case 0, 1 'numero de factura
            PonerFormatoEntero txtNumFac(Index)
    End Select
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

End Sub


Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtImporte_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtImporte(Index).Text = UCase(Trim(txtImporte(Index).Text))
    
    Select Case Index
        Case 0, 1 'importe de vencimiento
            PonerFormatoEntero txtImporte(Index)
    End Select
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

End Sub


Private Sub DividiVencimentosPorEntidadBancaria()

'    Set miRsAux = New ADODB.Recordset
'
'    Conn.Execute "DELETE FROM tmp347 WHERE codusu = " & vUsu.Codigo
'    '                                                               POR SI TUVIERAN MISMO BANCO, <> cta contable
'    NumeroDocumento = "select oficina,entidad from ctabancaria where not sufijoem is null "
'    NumeroDocumento = NumeroDocumento & " and entidad >0  and codmacta<>'" & Me.txtCta(3).Text & "' group by 1,2"
'    miRsAux.Open NumeroDocumento, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    NumeroDocumento = ""
'    While Not miRsAux.EOF
'        NumeroDocumento = NumeroDocumento & ", (" & miRsAux!Entidad & ")"
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
'    If NumeroDocumento = "" Then
'        NumeroDocumento = "(-1,-1)"
'    Else
'        NumeroDocumento = Mid(NumeroDocumento, 2) 'quitamos la primera coma
'    End If
'
'    NumeroDocumento = " (codbanco) in (" & NumeroDocumento & ")"
'
'    'Agrupamos los vencimientos por entidad,oficina menos los del banco por defecto
'    CuentasCC = "select codbanco,sum(impvenci + coalesce(gastos,0)) " & SQL
'    CuentasCC = CuentasCC & " AND " & NumeroDocumento & " GROUP BY 1"
'    miRsAux.Open CuentasCC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not miRsAux.EOF
'        CuentasCC = "insert into `tmpcierre1` (`codusu`,`cta`,`nomcta`,`acumPerD`) VALUES (" & vUsu.Codigo & ","
'        CuentasCC = CuentasCC & miRsAux.Fields(0) & ",0," & TransformaComasPuntos(CStr(miRsAux.Fields(1))) & ")"
'        Conn.Execute CuentasCC
'
'         miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
'    'Los del banco por defecto, y lo que no tenemos banco, es decir, el resto
'    '------------------------------------------------------------------------------
'    CuentasCC = SQL & " AND NOT " & NumeroDocumento & " GROUP BY 1,2"
'    'Vere la entidad y la oficina del PPAL
'    NumeroDocumento = DevuelveDesdeBD("concat(entidad,',',oficina)", "ctabancaria", "codmacta", txtCta(3).Text, "T")
'    NumeroDocumento = "Select " & NumeroDocumento & ",sum(impvenci + coalesce(gastos,0)) " & CuentasCC
'    miRsAux.Open NumeroDocumento, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not miRsAux.EOF
'        CuentasCC = "insert into `tmpcierre1` (`codusu`,`cta`,`nomcta`,`acumPerD`) VALUES (" & vUsu.Codigo & ","
'        CuentasCC = CuentasCC & miRsAux.Fields(0) & "," & miRsAux.Fields(1) & "," & TransformaComasPuntos(CStr(miRsAux.Fields(2))) & ")"
'        Conn.Execute CuentasCC
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'    espera 1
'
'
'    'Pongo codmacta y nombanco como corresponde
'    CuentasCC = "Select * from tmpcierre1 where codusu =" & vUsu.Codigo
'    miRsAux.Open CuentasCC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not miRsAux.EOF
'        NumeroDocumento = "nommacta"
'        CuentasCC = "ctabancaria.codmacta=cuentas.codmacta AND ctabancaria.entidad = " & miRsAux!Cta & " AND 1 "    'ctabancaria.oficina "
'        CuentasCC = DevuelveDesdeBD("ctabancaria.codmacta", "ctabancaria,cuentas", CuentasCC, "1", "N", NumeroDocumento)  'miRsAux!nomcta
'        If CuentasCC <> "" Then
'            CuentasCC = "UPDATE tmpcierre1 SET cta = '" & CuentasCC & "',nomcta ='" & DevNombreSQL(NumeroDocumento)
'            CuentasCC = CuentasCC & "' WHERE Cta = " & miRsAux!Cta & " AND nomcta =" & miRsAux!nomcta
'            Conn.Execute CuentasCC
'
'        End If
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
'    'Por si quiere borrar alguno de los repartios que hace
'    'Por si casao luego BORRAN la remesa a generar para ese banco, es decir , no uqieren llevarlo ahora
'    CuentasCC = "insert into tmp347(codusu,cta) select codusu,cta from tmpcierre1 WHERE codusu =" & vUsu.Codigo
'    Conn.Execute CuentasCC
'
'eDividir:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, Err.Description
'    End If
'    NumeroDocumento = ""
'    CuentasCC = ""
'    Set miRsAux = Nothing
'    Set RS = Nothing
End Sub

Private Function VencimientosPorEntidadBancaria() As String
Dim SQL As String

    VencimientosPorEntidadBancaria = ""

    SQL = " and length(cobros.iban) <> 0 and mid(cobros.iban,5,4) = (select mid(iban,5,4) from bancos where codmacta = " & DBSet(txtCuentas(2).Text, "T") & ")"
    
    VencimientosPorEntidadBancaria = SQL

End Function


Private Function GenerarRemesa(Opcion As Integer) As Boolean
Dim C As String
Dim NumeroRemesa As Long
Dim Rs As ADODB.Recordset
Dim J As Integer
Dim i As Integer
Dim ImporteQueda As Currency

    On Error GoTo eGenerarRemesa
    
    GenerarRemesa = False
    
    'Lo qu vamos a hacer es, primero bloquear la opcion de remesar
    If Opcion = 0 Then
        If Not BloqueoManual(True, "Remesas", "Remesas") Then
            MsgBox "Otro usuario esta generando remesas", vbExclamation
            Exit Function
        End If
    End If
    
    J = 0
    For i = 1 To lw1.ListItems.Count
        If lw1.ListItems(i).Checked Then
            J = 1
            Exit For
        End If
    Next i
    If J = 0 Then
        MsgBox "No se ha seleccionado cobros. Revise.", vbExclamation
        If Opcion = 0 Then BloqueoManual False, "Remesas", ""
        Exit Function
    End If
    
    
    'A partir de la fecha generemos leemos k remesa corresponde
    SQL = "select max(codigo) from remesas where anyo=" & Year(CDate(txtFecha(4).Text))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    NumeroRemesa = 0
    If Not miRsAux.EOF Then
        NumeroRemesa = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    NumeroRemesa = NumeroRemesa + 1
    
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select * from remesas where codigo=" & NumeroRemesa
    Cad = Cad & " AND anyo =" & Year(CDate(txtFecha(4).Text))
    Cad = Cad & " AND tiporem = 1"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then Cad = ""
    miRsAux.Close
    
    If Opcion = 0 Then
'        If Cad <> "" Then
'            MsgBox "Ya existe el codigo de remesa.", vbExclamation
'            Exit Function
'        End If
        Conn.BeginTrans
    Else
        If Cad = "" Then
            MsgBox "Error leyendo la remesa", vbExclamation
            Exit Function
        End If
    End If
    
    If Opcion = 0 Then
        'Ahora insertamos la remesa
        Cad = "INSERT INTO remesas (codigo, anyo, fecremesa,situacion,codmacta,descripcion,tiporem) "
        Cad = Cad & " VALUES ( "
        Cad = Cad & NumeroRemesa & "," & Year(CDate(txtFecha(4).Text)) & ",'" & Format(txtFecha(4).Text, FormatoFecha) & "','A','"
        Cad = Cad & txtCuentas(2).Text & "',' " & DevNombreSQL(txtRemesa.Text) & "',1)"
        Conn.Execute Cad
        
    Else
        'Paso la remesa a estado: A
        'Vuelvo a poner los vecnimientos a NULL para poder
        'meterlos luego
        
        '---remesa estado A
        Cad = "UPDATE Remesas SET Situacion = 'A'"
        Cad = Cad & ", descripcion ='" & DevNombreSQL(txtRemesa.Text) & "'"
        Cad = Cad & " WHERE codigo=" & NumeroRemesa
        Cad = Cad & " AND anyo =" & Year(CDate(txtFecha(4).Text))
        If Not Ejecuta(Cad) Then Exit Function
        
        Cad = "UPDATE cobros SET siturem=NULL, codrem=NULL, anyorem=NULL ,tiporem =NULL "
        Cad = Cad & " ,fecultco=NULL,ctabanc2=null, impcobro = NULL "
        Cad = Cad & " WHERE codrem = " & NumeroRemesa
        Cad = Cad & " AND anyorem=" & Year(CDate(txtFecha(4).Text)) & " AND tiporem = 1"
        If Not Ejecuta(Cad) Then Exit Function
    End If
    
    
    
    'Ahora cambiamos los cobros y les ponemos la remesa
    Cad = "UPDATE  cobros SET siturem= 'A',codrem= " & NumeroRemesa & ", anyorem =" & Year(CDate(txtFecha(4).Text)) & ","
    Cad = Cad & " tiporem = 1"
    'Lo cobrado tb a NULL
    
    'ponemos la cuenta de banco donde va remesado
'        Cad = Cad & ", ctabanc2 ='"
    
    
    
    'Para cada cobro UPDATE
    For J = 1 To lwCobros.ListItems.Count
       ' WHERE numserie='A' AND codfaccl=99588 AND fecfaccl='2003-12-01' AND numorden=4"
       With lwCobros.ListItems(J)
            If .Checked Then
                C = " WHERE numserie = '" & .Text & "' and numfactu = "
                C = C & Val(.SubItems(1)) & " and fecfactu ='" & Format(.SubItems(2), FormatoFecha)
                C = C & "' AND numorden =" & .SubItems(3)
                
                C = Cad & C
                Conn.Execute C
            End If
       End With
       
    Next J
    espera 0.5
    
    
    'Hacemos un select sum para el importe
    Cad = "Select sum(impvenci),sum(impcobro),sum(gastos) from cobros "
    Cad = Cad & " WHERE codrem=" & NumeroRemesa
    Cad = Cad & " AND anyorem =" & Year(CDate(txtFecha(4).Text))
    Cad = Cad & " AND tiporem = 1"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = "0"
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
                           'Impvenci                               impcobro                      gastos
            ImporteQueda = DBLet(miRsAux.Fields(0), "N") - DBLet(miRsAux.Fields(1), "N") + DBLet(miRsAux.Fields(2), "N")
            C = TransformaComasPuntos(CStr(ImporteQueda))
        End If
    End If
    miRsAux.Close
    
    Cad = "UPDATE Remesas SET importe=" & C
    Cad = Cad & " WHERE codigo=" & NumeroRemesa
    Cad = Cad & " AND anyo =" & Year(CDate(txtFecha(4).Text))
    Cad = Cad & " AND tiporem = 1"
    Conn.Execute Cad
    
        
    Set miRsAux = Nothing
    
    GenerarRemesa = True
    Conn.CommitTrans
    If Opcion = 0 Then BloqueoManual False, "Remesas", "Remesas"
    Exit Function
    
    
eGenerarRemesa:
    Conn.RollbackTrans
    MuestraError Err.Number, "Generar remesa", Err.Description
    If Opcion = 0 Then BloqueoManual False, "Remesas", "Remesas"
End Function

