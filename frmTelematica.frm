VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTelematica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   450
   ClientTop       =   630
   ClientWidth     =   11760
   Icon            =   "frmTelematica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
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
      Left            =   9300
      TabIndex        =   20
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
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
      Index           =   0
      Left            =   10500
      TabIndex        =   21
      Top             =   8040
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10830
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameLegalizacion 
      Height          =   4305
      Left            =   60
      TabIndex        =   1
      Top             =   3630
      Width           =   11595
      Begin VB.Frame FrameAgrupacion 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   5790
         TabIndex        =   57
         Top             =   210
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox Text2 
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
            Left            =   3960
            TabIndex        =   60
            Text            =   "Text2"
            Top             =   360
            Width           =   765
         End
         Begin VB.TextBox Text2 
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
            Left            =   3960
            TabIndex        =   59
            Text            =   "Text2"
            Top             =   0
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Cuentas anuales"
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
            Index           =   2
            Left            =   2715
            TabIndex        =   62
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label Label3 
            Caption         =   "Diario"
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
            Left            =   2715
            TabIndex        =   61
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label Label3 
            Caption         =   "N� presentaci�n"
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
            Index           =   0
            Left            =   480
            TabIndex        =   58
            Top             =   210
            Width           =   1770
         End
      End
      Begin VB.CheckBox chkCompartivo 
         Caption         =   "Comparativo"
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
         Left            =   3120
         TabIndex        =   54
         Top             =   3150
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "Balance situaci�n"
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
         Index           =   8
         Left            =   360
         TabIndex        =   53
         Top             =   3630
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkAgrupar 
         Caption         =   "Agrupar libros"
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
         Left            =   3510
         TabIndex        =   52
         Top             =   420
         Width           =   2505
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "P�rdidas y ganancias"
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
         Index           =   7
         Left            =   360
         TabIndex        =   51
         Top             =   3150
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "Inventario final"
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
         Index           =   6
         Left            =   8880
         TabIndex        =   50
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2385
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
         Index           =   2
         Left            =   1950
         TabIndex        =   47
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   6390
         TabIndex        =   42
         Top             =   1620
         Width           =   4575
         Begin VB.OptionButton optBalsum 
            Caption         =   "Mensual"
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
            Index           =   3
            Left            =   960
            TabIndex        =   56
            Top             =   480
            Width           =   1155
         End
         Begin VB.OptionButton optBalsum 
            Caption         =   "Men. acumulada"
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
            Index           =   4
            Left            =   2340
            TabIndex        =   55
            Top             =   480
            Width           =   1995
         End
         Begin VB.OptionButton optBalsum 
            Caption         =   "Trim. acumulada"
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
            Index           =   2
            Left            =   2340
            TabIndex        =   49
            Top             =   120
            Width           =   1995
         End
         Begin VB.OptionButton optBalsum 
            Caption         =   "Anual"
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
            Left            =   0
            TabIndex        =   44
            Top             =   120
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optBalsum 
            Caption         =   "Trimestral"
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
            Left            =   960
            TabIndex        =   43
            Top             =   120
            Width           =   1305
         End
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   1275
         Left            =   6210
         MouseIcon       =   "frmTelematica.frx":000C
         TabIndex        =   30
         Top             =   2310
         Width           =   5115
         Begin VB.CheckBox Check2 
            Caption         =   "9� nivel"
            Height          =   210
            Index           =   9
            Left            =   480
            TabIndex        =   40
            Top             =   1680
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CheckBox Check2 
            Caption         =   "8� nivel"
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
            Index           =   8
            Left            =   3420
            TabIndex        =   39
            Top             =   960
            Width           =   1545
         End
         Begin VB.CheckBox Check2 
            Caption         =   "7� nivel"
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
            Index           =   7
            Left            =   1860
            TabIndex        =   38
            Top             =   960
            Width           =   1365
         End
         Begin VB.CheckBox Check2 
            Caption         =   "6� nivel"
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
            Index           =   6
            Left            =   240
            TabIndex        =   37
            Top             =   960
            Width           =   1515
         End
         Begin VB.CheckBox Check2 
            Caption         =   "5� nivel"
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
            Index           =   5
            Left            =   3420
            TabIndex        =   36
            Top             =   600
            Width           =   1545
         End
         Begin VB.CheckBox Check2 
            Caption         =   "4� nivel"
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
            Index           =   4
            Left            =   1860
            TabIndex        =   35
            Top             =   600
            Width           =   1425
         End
         Begin VB.CheckBox Check2 
            Caption         =   "3� nivel"
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
            Index           =   3
            Left            =   240
            TabIndex        =   34
            Top             =   600
            Width           =   1395
         End
         Begin VB.CheckBox Check2 
            Caption         =   "2� nivel"
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
            Index           =   2
            Left            =   3420
            TabIndex        =   33
            Top             =   240
            Width           =   1545
         End
         Begin VB.CheckBox Check2 
            Caption         =   "1er nivel"
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
            Left            =   1860
            TabIndex        =   32
            Top             =   240
            Width           =   1485
         End
         Begin VB.CheckBox Check2 
            Caption         =   "�ltimo:  "
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
            Index           =   10
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Value           =   1  'Checked
            Width           =   1485
         End
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "Facturas recibidas"
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
         Index           =   5
         Left            =   3120
         TabIndex        =   29
         Top             =   2550
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "Facturas emitidas"
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
         Index           =   4
         Left            =   360
         TabIndex        =   28
         Top             =   2550
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "Inventario inicial"
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
         Index           =   3
         Left            =   6300
         TabIndex        =   27
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "Libro Mayor"
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
         Index           =   2
         Left            =   360
         TabIndex        =   26
         Top             =   1950
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "Balance de sumas y saldos"
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
         Index           =   1
         Left            =   6270
         TabIndex        =   25
         Top             =   1350
         Value           =   1  'Checked
         Width           =   3645
      End
      Begin VB.CheckBox chkLibro 
         Caption         =   "Libro diario"
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
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   1350
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmTelematica.frx":015E
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmTelematica.frx":02A8
         Top             =   855
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Informe"
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
         Index           =   2
         Left            =   210
         TabIndex        =   48
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   1680
         Picture         =   "frmTelematica.frx":03F2
         Top             =   360
         Width           =   240
      End
      Begin VB.Shape Shape1 
         Height          =   3015
         Left            =   120
         Top             =   1170
         Width           =   11295
      End
   End
   Begin VB.Frame FrameCuentas 
      Height          =   3555
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   11625
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
         Left            =   6930
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   1440
         Width           =   4125
      End
      Begin VB.CheckBox chkLanzaCtas 
         Caption         =   "Lanzar programa registro mercantil"
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
         Left            =   5730
         TabIndex        =   23
         Top             =   360
         Width           =   3795
      End
      Begin VB.TextBox txtDatos 
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
         Index           =   6
         Left            =   1740
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   3000
         Width           =   4245
      End
      Begin VB.TextBox txtDatos 
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
         Left            =   240
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtDatos 
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
         Left            =   6930
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2250
         Width           =   4095
      End
      Begin VB.TextBox txtDatos 
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
         Left            =   6060
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2250
         Width           =   825
      End
      Begin VB.TextBox txtDatos 
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
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2250
         Width           =   5745
      End
      Begin VB.TextBox txtDatos 
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
         Left            =   1770
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1440
         Width           =   5115
      End
      Begin VB.TextBox txtDatos 
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
         Left            =   240
         TabIndex        =   6
         Tag             =   "NIF|T|N|||||||"
         Text            =   "Text1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtpath 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   630
         Width           =   9255
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
         Left            =   9600
         TabIndex        =   2
         Text            =   "31/12/2015"
         Top             =   630
         Width           =   1455
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   11160
         TabIndex        =   63
         Top             =   180
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
      Begin VB.Label lblDatos 
         Caption         =   "Nombre para el programa del R. Mercantil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   6930
         TabIndex        =   46
         Top             =   1170
         Width           =   4380
      End
      Begin VB.Label lblDatos 
         Caption         =   "Actividad principal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1770
         TabIndex        =   19
         Top             =   2730
         Width           =   2670
      End
      Begin VB.Label lblDatos 
         Caption         =   "Tel�fono"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   2730
         Width           =   1260
      End
      Begin VB.Label lblDatos 
         Caption         =   "Poblaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   6930
         TabIndex        =   15
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label lblDatos 
         Caption         =   "C.P."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   6060
         TabIndex        =   13
         Top             =   1980
         Width           =   630
      End
      Begin VB.Label lblDatos 
         Caption         =   "Domicilio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Label lblDatos 
         Caption         =   "Nombre empresa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1800
         TabIndex        =   9
         Top             =   1170
         Width           =   2700
      End
      Begin VB.Label lblDatos 
         Caption         =   "N.I.F."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1170
         Width           =   870
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3120
         Picture         =   "frmTelematica.frx":047D
         Top             =   690
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Programa registro mercantil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   330
         Width           =   3585
      End
      Begin VB.Label Label6 
         Caption         =   "Inicio ejercicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   9600
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   11070
         Picture         =   "frmTelematica.frx":6CCF
         Top             =   660
         Width           =   240
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Presentaci�n digital de las cuentas anuales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   5
      Left            =   210
      TabIndex        =   64
      Top             =   240
      Width           =   6285
   End
   Begin VB.Label lblIndicador 
      Caption         =   "Label3"
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
      Left            =   120
      TabIndex        =   22
      Top             =   8040
      Width           =   6285
   End
End
Attribute VB_Name = "frmTelematica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1308


Public opcion As Byte
    '0.- Presentacion cuentas
    '1.- Legalizacion libros

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private frmDiario As frmInfDiarioOficial
Private frmConsulta As frmConExtrList
Private frmFactCli As frmFacturasCliListado
Private frmFactPro As frmFacturasProListado
Private frmAsientos As frmAsientosHcoList
Private frmBalanSyS As frmInfBalSumSal
Private frmPyG As frmInfBalances
Private frmSit As frmInfBalances


Dim PrimeraVez As Boolean
Dim SQL As String
Dim Cad As String
Dim Cont As Integer
Private Contador As Byte



Private Sub chkAgrupar_Click()
    Me.FrameAgrupacion.Visible = Me.chkAgrupar.Value = 1
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
        Exit Sub
    End If
    
'--
'    If vUsu.Nivel > 2 Then
'        MsgBox "No tiene permisos", vbExclamation
'        Exit Sub
'    End If
    
    
    If txtpath.Text = "" Then
        If MsgBox("No tiene el programa de legalizaci�n del registro mercantil" & vbCrLf & "�Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    End If
    If Not DatosOK Then Exit Sub
    

        
    'Pregunta del timepo
    SQL = "Este proceso puede llevar mucho tiempo. " & vbCrLf & vbCrLf & "�Desea continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    If opcion = 0 Then
        HacerPresentacionCuentas
    Else
        HacerLegalizaLibros
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    Limpiar Me
    lblIndicador.Caption = ""
    Me.FrameLegalizacion.Visible = False
    PonerDatosEmpresa
    
    Select Case opcion
    Case 0
        Label2(5).Caption = "Presentaci�n telem�tica de las cuentas anuales"
        Label2(5).ForeColor = &H800000
        Caption = "Presentaci�n cuentas"
       
        Cont = 3300
    Case 1
'        Caption = "Legalizaci�n libros"
'        Label2(5).Caption = "Presentacion telem�tica de libros formato digital"
'        Label2(5).ForeColor = &H80&
        Caption = "Presentacion telem�tica de libros formato digital"
        Text3(0).Text = Format(DateAdd("yyyy", -1, vParam.fechaini), "dd/mm/yyyy")
        Text3(2).Text = Format(DateAdd("yyyy", -1, vParam.fechafin), "dd/mm/yyyy")
        Me.FrameLegalizacion.Visible = True
        PonerNiveles
        Cont = 8655 '7320
    End Select
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    
    
    
    PonerPath
    
'    'En cont tengo donde iran los botones y demas
'    Me.lblIndicador.Top = Cont
'
'    Me.Command1(0).Top = Cont
'    Me.Command1(1).Top = Cont
'    Cont = Cont + Me.Command1(0).Height + 520
'    Me.Height = Cont
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPpal.Enabled = True
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(Text3(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
    On Error GoTo E1
    cd1.CancelError = True
    cd1.DialogTitle = "Archivo EXE"
    cd1.ShowOpen
    If cd1.FileTitle <> "" Then
        SQL = UCase(cd1.FileTitle)
        If SQL <> "D2.EXE" And SQL <> "LEGALIA.EXE" Then
            MsgBox "No es el archivo EXE que se esperaba( D2.EXE o Legalia.EXE)", vbExclamation
        Else
            txtpath.Text = cd1.FileName
            NumRegElim = InStr(1, cd1.FileName, cd1.FileTitle)
            txtpath.Tag = Mid(cd1.FileName, 1, NumRegElim - 1)
        End If
    Else
        MsgBox "No es un archivo EXE", vbExclamation
    End If
    If txtpath.Text <> "" Then
        chkLanzaCtas.Value = 1
    Else
        chkLanzaCtas.Value = 0
    End If
        
    Exit Sub
E1:
    Err.Clear
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    Text3(0).Tag = Index
    If Text3(Index).Text <> "" Then
        frmC.Fecha = CDate(Text3(Index).Text)
    Else
        frmC.Fecha = Now
    End If
    frmC.Show vbModal
    Set frmC = Nothing
            
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For Cont = 0 To Me.chkLibro.Count - 1
        chkLibro(Cont).Value = Index
    Next
End Sub

'Private Sub Option1_Click(Index As Integer)
'    If Option1(1).Value Then
'        Combo1.Enabled = True
'    Else
'        Combo1.Enabled = False
'        Combo1.ListIndex = -1
'    End If
'End Sub


Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text <> "" Then
        If Not IsNumeric(Text2(Index).Text) Then Text2(Index).Text = ""
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    With Text3(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'++
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYFecha KeyAscii, 0
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    Image2_Click (indice)
End Sub
'++

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index).Text)
    If Text3(Index).Text <> "" Then
        If Not EsFechaOK(Text3(Index)) Then PonerFoco Text3(Index)
    End If
End Sub



Private Sub PonerFoco(ByRef Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub





Private Sub PonerDatosEmpresa()
    'Datos basicos
    txtDatos(1).Text = vEmpresa.nomempre
    Text3(0).Text = Format(vParam.fechaini, "dd/mm/yyyy")
    Text1.Text = vEmpresa.nomresum
    
    'Ponemos los datos
    Set miRsAux = New ADODB.Recordset
    SQL = "SELECT * from Empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    With miRsAux
        If Not miRsAux.EOF Then
            txtDatos(0).Text = DBLet(!nifempre)
            txtDatos(2).Text = Trim(DBLet(!siglasvia) & " " & DBLet(!Direccion))
            txtDatos(3).Text = DBLet(!codpos)
            txtDatos(4).Text = DBLet(!poblacion)
            txtDatos(5).Text = DBLet(!tfnocontacto)
        End If
        .Close
    End With
End Sub
        


Private Sub PonerPath()

    On Error GoTo EPonerPath
    txtpath.Text = ""   'EXE
    txtpath.Tag = ""    'CARPETA
    
    'Buscaremos el archivo en dos sitios, en Archivos de programa o en Program FIles
    Cad = "C:\Archivos de Programa\Adhoc\"
    If opcion = 0 Then
        CadenaDesdeOtroForm = "D2\"
        SQL = "D2.exe"
    Else
        CadenaDesdeOtroForm = "Legalia\"
        SQL = "Legalia.exe"
    End If
    
    
    If Dir(Cad & CadenaDesdeOtroForm & SQL, vbArchive) <> "" Then
        'Esta aqui el archivo
        txtpath.Text = Cad & CadenaDesdeOtroForm & SQL
        txtpath.Tag = Cad & CadenaDesdeOtroForm
    Else
        Cad = "C:\Program Files\Adhoc\"
        If Dir(Cad & CadenaDesdeOtroForm & SQL, vbArchive) <> "" Then
            txtpath.Text = Cad & CadenaDesdeOtroForm & SQL
            txtpath.Tag = Cad & CadenaDesdeOtroForm
            
        Else
            Cad = "C:\Program Files (x86)\Adhoc\"
            If Dir(Cad & CadenaDesdeOtroForm & SQL, vbArchive) <> "" Then
                txtpath.Text = Cad & CadenaDesdeOtroForm & SQL
                txtpath.Tag = Cad & CadenaDesdeOtroForm
            End If
        End If
    End If
    If txtpath.Text <> "" Then
        chkLanzaCtas.Value = 1
    Else
        chkLanzaCtas.Value = 0
    End If
    
    Exit Sub
EPonerPath:
    MuestraError Err.Number, "Poner PATH defecto" & vbCrLf & Cad
End Sub





Private Function DatosOK() As Boolean
    DatosOK = False

        '----------------------------------------------------------------
        '            Comunes
        '----------------------------------------------------------------
        If Not Comprobar_NIF(txtDatos(0).Text) Then
            If MsgBox("NIF incorrecto.  �Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        
        If Text3(0).Text = "" Then
            MsgBox "Fecha incio ejercicio incorrecta", vbExclamation
            Exit Function
        End If
        
        
        Cont = 0
        For NumRegElim = 1 To 5
            txtDatos(NumRegElim).Text = Trim(txtDatos(NumRegElim).Text)
            If txtDatos(NumRegElim).Text = "" Then Cont = Cont + 1
        Next NumRegElim
        
        If Cont > 0 Then
            SQL = "Existen campos sin rellenar. �Desea continuar igualmente?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
        

        
    If opcion = 0 Then
        'Solo Cuentas
        'Vemos si tiene los balances configurados 3 y 4  'Abreviados
        Cad = DevuelveDesdeBD("nombalan", "balances", "numbalan", 3)
        If Cad = "" Then
            MsgBox "Falta balance perdidas y ganancias abreviado "
            Exit Function
        End If
        
        Cad = DevuelveDesdeBD("nombalan", "balances", "numbalan", 4)
        If Cad = "" Then
            MsgBox "Falta balance situacion abreviado "
            Exit Function
        End If
        
    Else
        '----------------------------------------------------------------
        '           SOLO Legalizacion de libros
        '----------------------------------------------------------------
        
       
        Cont = 0
        For NumRegElim = 0 To Me.chkLibro.Count - 1
            If chkLibro(NumRegElim).Value = 1 Then Cont = Cont + 1
        Next NumRegElim
        If Cont = 0 Then
            MsgBox "Seleccione algun libro para legalizar", vbExclamation
            Exit Function
        End If
        
        If Text3(2).Text = "" Then
            MsgBox "Fecha informes incorrecta", vbExclamation
            Exit Function
        End If
        'Ahora veremos, si ha marcado diario resumen los digitos
        If chkLibro(0).Value = 1 And Combo1.ListIndex = -1 Then
            'Diario resumen. Comprobar digitos
            MsgBox "Seleccione un nivel", vbExclamation
            Exit Function
        End If
        
        
        'Si presnta balances
        If chkLibro(1).Value = 1 Then
            Cont = 0
            For NumRegElim = 1 To 10
                If Check2(NumRegElim).Value = 1 Then Cont = Cont + 1
            Next NumRegElim
            
'            If optBalsum(3).Value = True Or optBalsum(4).Value = True Then
'                If Cont > 1 Then
'                    MsgBox "Mensual como m�ximo un  nivel para el balance de sumas y saldos", vbExclamation
'                    Exit Function
'                 End If
'            Else
'                If Cont > 2 Then
'                    'Dos pq 4 trimestres como maximo a dos digitos que guarda el listado 31..39 tenemos suficionete.
'                    MsgBox "Trimestral como m�ximo dos niveles para el balance de sumas y saldos", vbExclamation
'                    Exit Function
'                End If
'            End If
        End If
        
        'Si esta agrupar, entonces tiene k existir el archivo
        If Me.chkAgrupar.Value = 1 Then
            If Dir(App.Path & "\pdftk.exe", vbArchive) = "" Then
                MsgBox "No tiene el archivo que falta para la legalizacion de los libros", vbExclamation
                Exit Function
            End If
            
            If Not CompruebarCarpetaAgrupacion Then Exit Function
        
        End If
        
        
        
        
      'Si esta marcado balances, y son acumulados obligaremos a mrcar la agrupacion
      'ya que si no el programa de LEGALIA dara erroes al solaparse las fechas
      If chkLibro(1).Value Then
        If Me.optBalsum(2).Value Or Me.optBalsum(4).Value Then
            If Me.chkAgrupar.Value = 0 Then
                MsgBox "Si selecciona balances acumulados deber� agrupar los libros.", vbExclamation
                Exit Function
            End If
        End If
      End If
        
    End If
    DatosOK = True
End Function

Private Function CompruebarCarpetaAgrupacion() As Boolean
    On Error GoTo E1
    CompruebarCarpetaAgrupacion = False
    
    If Dir(App.Path & "\temp", vbDirectory) = "" Then
        MkDir App.Path & "\temp"
    Else
        If Dir(App.Path & "\temp\*.*", vbArchive) <> "" Then Kill App.Path & "\temp\*.*"
    End If
    CompruebarCarpetaAgrupacion = True
    Exit Function
E1:
    MuestraError Err.Number, "Comprueba carpeta Agrupar"
End Function


Private Sub HacerPresentacionCuentas()

    On Error GoTo EH

    'Crearemos la carpeta DATA
    If Dir(txtpath.Tag & "Data", vbDirectory) = "" Then MkDir txtpath.Tag & "Data"
    
    'Crearemos lo de empresa
    SQL = txtpath.Tag & "Data\" & vEmpresa.nomresum & ".AE"   'Abreviada en euros
    
    If Dir(SQL, vbDirectory) <> "" Then
        Cad = "Ya existen datos para la empresa: " & vEmpresa.nomresum & ".   �Desea continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Kill SQL & "\*.*"
    Else
        MkDir SQL
    End If
    
    lblIndicador.Caption = ""
    Me.Refresh
    If GenerarDatosFicheros(SQL) Then
        
        'Si esta marcado se lanzara el programa
        If Me.chkLanzaCtas.Value = 1 Then
            Shell txtpath.Text, vbMaximizedFocus
        Else
            SQL = "Ya ha generado los datos para presentaci�n de las cuentas anuales." & vbCrLf
            SQL = SQL & "Cuando desee ejecute el programa: " & vbCrLf & "   " & txtpath.Text & vbCrLf
            SQL = SQL & "y seleccione su empresa: " & vEmpresa.nomresum
            MsgBox SQL, vbInformation
            
        End If
    End If
    
    Exit Sub
EH:
    MuestraError Err.Number, "hacer presentacion"
End Sub


Private Function GenerarDatosFicheros(ByVal vPath As String) As Boolean

'        'Cambiar numfac prov
        Dim Rs As ADODB.Recordset
        Dim Cad2 As String
        Dim Valor As Currency
        Dim Anyo As String
        
        
        On Error GoTo EGenerarDatosFicheros
        
        Set Rs = New ADODB.Recordset
        
        GenerarDatosFicheros = False
        
        '-----------------------------------------------
        ' Fichero Descripcion
        Me.lblIndicador.Caption = "Descripcion"
        Me.Refresh
        Cont = FreeFile
        Open vPath & "\DESC.TXT" For Output As #Cont
        SQL = vEmpresa.nomempre & ". Generado: " & Format(Now, "dd/mm/yyyy hh:mm")
        Print #Cont, SQL
        Close #Cont
        
        
        'Fichero con los FIcheros k van
        Me.lblIndicador.Caption = "Fichero"
        Me.Refresh
        Cont = FreeFile
        Open vPath & "\FICHERO.TXT" For Output As #Cont
        Print #Cont, "FICHERO.TXT"
        Print #Cont, "DATOS.ASC"
        Close #Cont
                
        'Los datos
        Me.lblIndicador.Caption = "Datos"
        Me.Refresh
        
        Cont = FreeFile
        Open vPath & "\DATOS.ASC" For Output As #Cont
        Anyo = Year(CDate(Text3(0).Text))
        
        'Cabecera
        SQL = "1"
        SQL = SQL & txtDatos(0).Text
        SQL = SQL & Anyo
        SQL = SQL & "000000"
        SQL = SQL & "02"
        SQL = SQL & Space(8)
        SQL = SQL & Space(8)
        SQL = SQL & Mid(txtDatos(3) & "  ", 1, 2)   'Dos digitos cod pos
        
        'Nombre Empresa
        Cad = txtDatos(1).Text
        Cad = Mid(Cad & Space(50), 1, 50)
        SQL = SQL & Cad
        'Domicilia
        Cad = txtDatos(2).Text
        Cad = Mid(Cad & Space(40), 1, 40)
        SQL = SQL & Cad
        'Municipio
        Cad = txtDatos(4).Text
        Cad = Mid(Cad & Space(30), 1, 30)
        SQL = SQL & Cad
        
        SQL = SQL & Mid(txtDatos(3).Text & "     ", 1, 5)
        
        SQL = SQL & " " & Mid(txtDatos(5).Text & "  ", 1, 2)
        SQL = SQL & Mid(txtDatos(5).Text & "       ", 1, 7)
        SQL = SQL & "000"
        SQL = SQL & "000"
        'Actividad principal
        Cad = ""
        Cad = Mid(Cad & Space(80), 1, 80)
        SQL = SQL & Cad

        SQL = SQL & "0000000000"
        SQL = SQL & "00000"

        Print #Cont, SQL
    
    
        'Balance de sitaucion
        '----------------------------------
        '----------------------------------
        Me.lblIndicador.Caption = "Situacion"
        Me.Refresh
        GeneraDatosBalanceConfigurable 4, 12, 2003, 12, 2002, True, -1
        Me.lblIndicador.Caption = "Escribir resultados 1"
        Me.Refresh
        espera 1
        SQL = "select *  from Usuarios.ztmpimpbalan WHERE codusu=" & vUsu.Codigo & " AND not (libroCD is null)"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad2 = "00000000000000"
        While Not Rs.EOF
            SQL = "2"
            SQL = SQL & txtDatos(0).Text
            SQL = SQL & Anyo
            SQL = SQL & Rs!LibroCD
            
            'Importe1
            If IsNull(Rs!Importe1) Then
                Valor = 0
            Else
                Valor = Rs!Importe1
            End If
            Valor = Valor * 100
            Cad = CStr(Abs(Valor))
            If Valor > 0 Then
                SQL = SQL & "+"
            Else
                If Valor < 0 Then
                    SQL = SQL & "-"
                Else
                    SQL = SQL & " "
                End If
            End If
            Cad = Right(Cad2 & Cad, 15)
            SQL = SQL & Cad
            
            'Importe anterior
            'Importe1
            If IsNull(Rs!importe2) Then
                Valor = 0
            Else
                Valor = Rs!importe2
            End If
            Valor = Valor * 100
            Cad = CStr(Abs(Valor))
            If Valor > 0 Then
                SQL = SQL & "+"
            Else
                If Valor < 0 Then
                    SQL = SQL & "-"
                Else
                    SQL = SQL & " "
                End If
            End If
            Cad = Right(Cad2 & Cad, 15)
            SQL = SQL & Cad
            Print #Cont, SQL
            
            'Siguiente
            Rs.MoveNext
        Wend
        Rs.Close
        espera 1
        
        '^Perdidas y ganancias
        '----------------------------------
        '----------------------------------
        Me.lblIndicador.Caption = "Perdidas"
        Me.Refresh
        GeneraDatosBalanceConfigurable 3, 12, 2003, 12, 2002, True, -1
        Me.lblIndicador.Caption = "Escribir resultados 2"
        Me.Refresh
        espera 1

        SQL = "select *  from Usuarios.ztmpimpbalan WHERE codusu=" & vUsu.Codigo & " AND not (libroCD is null)"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad2 = "00000000000000"
        While Not Rs.EOF
            SQL = "2"
            SQL = SQL & txtDatos(0).Text
            SQL = SQL & Anyo
            SQL = SQL & Rs!LibroCD
            
            'Importe1
            If IsNull(Rs!Importe1) Then
                Valor = 0
            Else
                Valor = Rs!Importe1
            End If
            Valor = Valor * 100
            Cad = CStr(Abs(Valor))
            If Valor > 0 Then
                SQL = SQL & "+"
            Else
                If Valor < 0 Then
                    SQL = SQL & "-"
                Else
                    SQL = SQL & " "
                End If
            End If
            Cad = Right(Cad2 & Cad, 15)
            SQL = SQL & Cad
            
            'Importe anterior
            'Importe2
            If IsNull(Rs!importe2) Then
                Valor = 0
            Else
                Valor = Rs!importe2
            End If
            Valor = Valor * 100
            Cad = CStr(Abs(Valor))
            If Valor > 0 Then
                SQL = SQL & "+"
            Else
                If Valor < 0 Then
                    SQL = SQL & "-"
                Else
                    SQL = SQL & " "
                End If
            End If
            Cad = Right(Cad2 & Cad, 15)
            SQL = SQL & Cad
            Print #Cont, SQL
            
            'Siguiente
            Rs.MoveNext
        Wend
        Rs.Close
        'Fi fichreo
        Close #Cont
        
        
        GenerarDatosFicheros = True
        
    
       Exit Function
        
EGenerarDatosFicheros:
    MuestraError Err.Number, "Generar Datos Ficheros"
        
End Function



Private Sub PonerNiveles()

    
    Combo1.AddItem "�ltimo"
    Combo1.ItemData(Combo1.NewIndex) = vEmpresa.DigitosUltimoNivel
    For NumRegElim = 1 To vEmpresa.numnivel - 1
        Cont = DigitosNivel(CInt(NumRegElim))
        Cad = "Digitos: " & Cont
        Check2(NumRegElim).Visible = True
        Check2(NumRegElim).Caption = Cad
        Check2(NumRegElim).Tag = Cont
        
                
        Combo1.AddItem "Nivel :   " & NumRegElim
        Combo1.ItemData(Combo1.NewIndex) = Cont
    Next NumRegElim
    For Cont = NumRegElim To 9
        Check2(Cont).Visible = False
    Next Cont
End Sub


'----------------------------------------------------------------
'-----------------------------------------------------------------

Private Sub HacerLegalizaLibros()

    'Iremos uno a uno generando los libros k haya pedido
    
    'Crearemos la carpeta DATA
    If Dir(txtpath.Tag & "Data", vbDirectory) = "" Then MkDir txtpath.Tag & "Data"
    
    'Crearemos lo de empresa
    SQL = txtpath.Tag & "Data\" & vEmpresa.nomresum    'Abreviada en euros
    
    If Dir(SQL, vbDirectory) <> "" Then
        Cad = "Ya existen datos para la empresa: " & vEmpresa.nomresum & ".   �Desea continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        If Dir(SQL & "\*.*", vbArchive) <> "" Then Kill SQL & "\*.*"
    Else
        MkDir SQL
    End If
  
    
    frmPpal.Hide
    Me.Refresh
    
    'Borramos todos los registros de businmov
    Cad = "Delete from tmpfaclin WHERE codusu= " & vUsu.Codigo
    Conn.Execute Cad
    espera 1
    Me.Enabled = False
    If GenerarLibrosLegaliza(SQL) Then
             'Si esta marcado se lanzara el programa
        If Me.chkLanzaCtas.Value = 1 Then
            Shell txtpath.Text, vbMaximizedFocus
            espera 2
        Else
            SQL = "Ya ha generado los datos para presentaci�n de las cuentas anuales." & vbCrLf
            SQL = SQL & "Cuando desee ejecute el programa: " & vbCrLf & "   " & txtpath.Text & vbCrLf
            SQL = SQL & "y seleccione su empresa: " & vEmpresa.nomresum
            MsgBox SQL, vbInformation
        End If
        
        
    End If
    frmPpal.Show
    Me.Enabled = True
    Me.lblIndicador.Caption = ""
    Me.Refresh
    Me.SetFocus
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
End Sub

Private Sub CursorReloj()
    DoEvents
    Me.Refresh
    Me.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
End Sub


Private Function GenerarLibrosLegaliza(ByVal vPath As String) As Boolean
Dim CadenaLegaliza As String
Dim F As Date
Dim F2 As Date
Dim i As Integer
Dim J As Integer
Dim Incremento As Integer
Dim Bucle As Byte
Dim NumeroPresentacion As Integer
    
    
    
    GenerarLibrosLegaliza = False

    'Fichero con la descripcion de la empresa
    Me.lblIndicador.Caption = "Descripcion"
    Me.Refresh
    Cont = FreeFile
    Open vPath & "\DESC.TXT" For Output As #Cont
    SQL = vEmpresa.nomempre & ". Generado: " & Format(Now, "dd/mm/yyyy hh:mm")
    Print #Cont, SQL
    Close #Cont

    CursorReloj
    
    'Cadena legalizacion
    F = CDate(Text3(0).Text)
    CadenaLegaliza = Text3(2).Text & "|" & Format(F, "dd/mm/yyyy") & "|"
    F2 = DateAdd("yyyy", 1, F)  'mas 1 a�o
    F2 = DateAdd("d", -1, F2) 'menos un dia
    CadenaLegaliza = CadenaLegaliza & Format(F2, "dd/mm/yyyy") & "|"
    
    Contador = 0

    'DIARIO
    If Me.chkLibro(0).Value = 1 Then
    
        lblIndicador.Caption = "Generar Diario oficial"
        lblIndicador.Refresh
        CadenaDesdeOtroForm = ""
        SQL = CadenaLegaliza
        
        'Genero el resumido
        SQL = SQL & Me.Combo1.ItemData(Combo1.ListIndex) & "|"
        
        Set frmDiario = New frmInfDiarioOficial

        frmDiario.Legalizacion = SQL
        frmDiario.Show vbModal
        
        Set frmDiario = Nothing
            
        If CadenaDesdeOtroForm = "" Then
            'ERror
            GoTo Salida
        Else
            'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
            NumeroPresentacion = 1
            If Me.chkAgrupar.Value Then
                If Val(Text2(0).Text) > 1 Then NumeroPresentacion = Text2(0).Text
            End If
            If Not CopiarArchivoLega(vPath, 0, NumeroPresentacion, F, F2) Then GoTo Salida
        End If
    End If


    CursorReloj
    'LIBRO MAYOR o CONSULTA DE EXTRACTOS
    If Me.chkLibro(2).Value = 1 Then
    
        lblIndicador.Caption = "Generar libro mayor"
        lblIndicador.Refresh
        CadenaDesdeOtroForm = ""
        SQL = CadenaLegaliza
        
        Set frmConsulta = New frmConExtrList
        frmConsulta.Legalizacion = SQL
        frmConsulta.Show vbModal
        Set frmConsulta = Nothing
        
        If CadenaDesdeOtroForm = "" Then
            'ERror
            GoTo Salida
        Else
            'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
            If Not CopiarArchivoLega(vPath, 1, 1, F, F2) Then GoTo Salida
        End If
    End If
    
    
    
    CursorReloj
    'INVEMTARIO INCIAL
    If Me.chkLibro(3).Value = 1 Then
    
        lblIndicador.Caption = "Generar inventario inicial"
        lblIndicador.Refresh
        CadenaDesdeOtroForm = ""
        SQL = CadenaLegaliza
        
        Set frmAsientos = New frmAsientosHcoList
        
        frmAsientos.Legalizacion = SQL & "1|"
        frmAsientos.Show vbModal
        
        Set frmAsientos = Nothing
        
'        frmListado.opcion = 35
'        frmListado.Legalizacion = SQL
'        frmListado.EjerciciosCerrados = False   'Habra k comprobarlo
'        frmListado.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            'ERror
             GoTo Salida
        Else
            
            'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
            'Antes enviamos f y f2
            'If Not CopiarArchivoLega(vPath, 2, 1, F, F2) Then GoTo Salida
            'Ahora enviamos F y F, para que no se solpane las fechascon el invetario fial al cierre
            If Not CopiarArchivoLega(vPath, 2, 1, F, F) Then GoTo Salida
        End If
        espera 0.5
    End If
    



    CursorReloj
    'Facturas emitidas
    If Me.chkLibro(4).Value = 1 Then
    
        lblIndicador.Caption = "Facturas emitidas"
        lblIndicador.Refresh
        CadenaDesdeOtroForm = ""
        SQL = CadenaLegaliza
        
        Set frmFactCli = New frmFacturasCliListado
        
        frmFactCli.Legalizacion = SQL
        frmFactCli.Show vbModal
        
        Set frmFactCli = Nothing
'        frmListado.opcion = 37
'        frmListado.Legalizacion = SQL
'        frmListado.EjerciciosCerrados = False   'Habra k comprobarlo
'        frmListado.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            'ERror
            GoTo Salida
        Else
            'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
            If Not CopiarArchivoLega(vPath, 4, 1, F, F2) Then GoTo Salida
        End If
    End If


    CursorReloj
    'Facturas recibidas
    If Me.chkLibro(5).Value = 1 Then
    
        lblIndicador.Caption = "Facturas recibidas"
        lblIndicador.Refresh
        CadenaDesdeOtroForm = ""
        SQL = CadenaLegaliza

        Set frmFactPro = New frmFacturasProListado
        
        frmFactPro.Legalizacion = SQL
        frmFactPro.Show vbModal
        
        Set frmFactPro = Nothing


'        frmListado.opcion = 38
'        frmListado.Legalizacion = SQL
'        frmListado.EjerciciosCerrados = False   'Habra k comprobarlo
'        frmListado.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            'ERror
            GoTo Salida
        Else
            'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
            If Not CopiarArchivoLega(vPath, 5, 1, F, F2) Then GoTo Salida
        End If
    End If




'Balance de situacion, perdidas y ganacias
  '------------------------------------------
  CursorReloj
  
  If Me.chkLibro(7).Value = 1 Then
  
        lblIndicador.Caption = "Balance perdidas y ganancias"
        lblIndicador.Refresh
        CadenaDesdeOtroForm = ""
        SQL = CadenaLegaliza
        
        Set frmPyG = New frmInfBalances
        frmPyG.opcion = 1
        frmPyG.Legalizacion = SQL & Me.chkCompartivo.Value & "|"
        frmPyG.Show vbModal
        
        Set frmPyG = Nothing
        
'        frmListado.opcion = 39
'        frmListado.Legalizacion = SQL & Me.chkCompartivo.Value & "|"
'        frmListado.EjerciciosCerrados = False   'Habra k comprobarlo
'        frmListado.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            'ERror
            GoTo Salida
        Else
            'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
            If Not CopiarArchivoLega(vPath, 6, 1, F, F2) Then GoTo Salida
        End If
   End If
        
        
        
  'Situacion
  If Me.chkLibro(8).Value = 1 Then
        
        lblIndicador.Caption = "Balance situaci�n"
        lblIndicador.Refresh
        CadenaDesdeOtroForm = ""
        SQL = CadenaLegaliza

        Set frmSit = New frmInfBalances
        
        frmSit.opcion = 0
        frmSit.Legalizacion = SQL & Me.chkCompartivo.Value & "|"
        frmSit.Show vbModal
        
        Set frmSit = Nothing


'        frmListado.opcion = 40
'        frmListado.Legalizacion = SQL & chkCompartivo.Value & "|"
'        frmListado.EjerciciosCerrados = False   'Habra k comprobarlo
'        frmListado.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            'ERror
            GoTo Salida
        Else
            'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
            If Not CopiarArchivoLega(vPath, 7, 1, F, F2) Then GoTo Salida
        End If
        
        
        
        
   End If

  
  
  
  'INVENTARIO final cierre
  If Me.chkLibro(6).Value = 1 Then
        

        lblIndicador.Caption = "Inventario final cierre"
        lblIndicador.Refresh
        CadenaDesdeOtroForm = ""
        SQL = CadenaLegaliza
        
        Set frmBalanSyS = New frmInfBalSumSal
        frmBalanSyS.Legalizacion = SQL & "10|1|"
        frmBalanSyS.Show vbModal
        Set frmBalanSyS = Nothing
        
'        frmListado.opcion = 41
'        frmListado.Legalizacion = SQL & "10|"
'        frmListado.EjerciciosCerrados = False   'Habra k comprobarlo
'        frmListado.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            'ERror
            GoTo Salida
        Else
            'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
            'Si esta marcado el inventario incial entonce el libros sera el 2
            If Me.chkLibro(3).Value = 1 Then
                SQL = "2"
            Else
                SQL = "1"
            End If
            If Not CopiarArchivoLega(vPath, 8, CInt(SQL), F, F2) Then GoTo Salida
        End If
        
        
        
        
   End If
  
  
    







    'El ulttimo es el de sumas y saldos pq modifica las fechas

    'Balance de sumas y saldos
    CursorReloj
    If Me.chkLibro(1).Value = 1 Then
        lblIndicador.Caption = "Generar balance sumas y saldos"
        lblIndicador.Refresh
        
        If optBalsum(0).Value Then
            NumRegElim = 1
        Else
            'TRIMESTRAL o mensual
            If optBalsum(1).Value Or optBalsum(2).Value Then
                NumRegElim = 4
                Incremento = 3
            Else
                NumRegElim = 12
                Incremento = 1
            End If
            
            
        End If

        J = 1 'Contador de libros
        'Para cada periodo(si es anual sera uno solo)
        F = CDate(Text3(0).Text)
        For Bucle = 1 To NumRegElim
            
            If NumRegElim > 1 Then
                CadenaLegaliza = Text3(2).Text & "|" & Format(F, "dd/mm/yyyy") & "|"
                F2 = DateAdd("m", Incremento, F)  'mas 3 meses
                F2 = DateAdd("d", -1, F2) 'menos un dia
                CadenaLegaliza = CadenaLegaliza & Format(F2, "dd/mm/yyyy") & "|"
            End If
            'Debug.Print CadenaLegaliza
            For i = 1 To 10
                If Check2(i).Value = 1 Then
                    'Para el nivle
                    Me.lblIndicador.Caption = "Balances. Fecha: " & F & "    Nivel: " & i
                    Me.lblIndicador.Refresh
                    CursorReloj
                    SQL = CadenaLegaliza
                    SQL = SQL & i & "|"
                    CadenaDesdeOtroForm = ""
                    
                    
                    Set frmBalanSyS = New frmInfBalSumSal
                    frmBalanSyS.Legalizacion = SQL & "0|"
                    frmBalanSyS.Show vbModal
                    Set frmBalanSyS = Nothing
                    
                    
'                    frmListado.opcion = 36
'                    frmListado.Legalizacion = SQL
'                    frmListado.EjerciciosCerrados = False   'Habra k comprobarlo
'                    frmListado.Show vbModal
                    If CadenaDesdeOtroForm = "" Then
                        'Error
                        GoTo Salida
                    Else
                        'OK. Copiamos el archivo donde corresponda, con el nombre k corresponda
                        If Not CopiarArchivoLega(vPath, 3, J, F, F2) Then GoTo Salida
                        J = J + 1
                    End If
                    Me.Refresh
                    espera 0.5
                End If  'Del check2(I)
            Next i
            'Actualizamos fecha
            If NumRegElim > 1 Then
                If optBalsum(1).Value Or optBalsum(3).Value Then
                    'Trimestral o mensual  normal
                    F = DateAdd("d", 1, F2) 'mas un dia
                Else
                    
                    If optBalsum(2).Value Then
                        'Trimestral acumulado
                        Incremento = Incremento + 3
                    Else
                        'MENSUAL acumulado
                        Incremento = Incremento + 1
                    End If
                End If
            End If
            
        Next Bucle
        
   End If
        
        
        
   
    
    
   If Me.chkAgrupar.Value Then
        lblIndicador.Caption = "Unificando libros"
        DoEvents
        Me.Refresh
        espera 1
        
        F = CDate(Text3(0).Text)
        'F2 = DateAdd("yyyy", 1, F)  'mas 1 a�o
        'F2 = DateAdd("d", -1, F2) 'menos un dia
   
        'Creamos la TAPA
        CrearSeparadorTapa False, Cad, 0
   
        If Dir(App.Path & "\Docum.pdf", vbArchive) <> "" Then
            Kill App.Path & "\Docum.pdf"
            Me.Refresh
            espera 0.5
        End If
   
        If Dir(App.Path & "\Docum2.pdf", vbArchive) <> "" Then
            Kill App.Path & "\Docum2.pdf"
            Me.Refresh
            espera 0.5
        End If
   
   
        'lanzar SHELL agrupacion
        
        'cad = """" & App.path & "\pdftk.exe"" """ & App.path & "\temp\*.pdf"" cat output """ & App.path & "\Docum2.pdf"" verbose"

        MontaCadenaExe
        If Cad = "" Then GoTo Salida
        Shell Cad, vbMaximizedFocus
        
        If Not CopiarDocum Then
            MsgBox "No se ha podido unificar los libros", vbExclamation
            GoTo Salida
        End If
        
        
        
        lblIndicador.Caption = "Llevando fichero generado"
        Me.Refresh
        espera 0.1
   
        'Metemos el libro en la carpeta de LEGALIA
        NumeroPresentacion = Val(Text2(1).Text)
        If NumeroPresentacion = 0 Then NumeroPresentacion = 1
        If Not CopiarArchivoLega(vPath, 100, NumeroPresentacion, F, F2) Then GoTo Salida
   End If
        
    Cad = vPath 'Fijo el path para el ultimo fichero
    CreaDatosTXT
        
    GenerarLibrosLegaliza = True
    
Salida:
    
End Function

Private Sub MontaCadenaExe()
Dim Fs, F
    Cad = ""
    SQL = ""
    On Error GoTo eMontaCadenaExe
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set F = Fs.GetFolder(App.Path)
    SQL = F.shortpath & "\pdftk.exe"
    SQL = SQL & " " & F.shortpath & "\temp\*.pdf cat output " & F.shortpath & "\Docum2.pdf verbose"
    Cad = SQL
    
eMontaCadenaExe:
    If Err.Number <> 0 Then MuestraError Err.Number, "Sub: MontaCadenaExe"
    Set Fs = Nothing
    Set F = Nothing
End Sub


'numero
Private Function CopiarArchivoLega(vPa As String, opcion As Byte, numero As Integer, F1 As Date, F2 As Date) As Boolean
Dim Insertar As Boolean
Dim i As Integer

    On Error GoTo eCopiarArchivoLega
    CopiarArchivoLega = False
'BAL_SUMS_002.PDF
'DIARIO_001.PDF
'FAC_EMIT_005.PDF
'FAC_RECI_006.PDF
'INVENTAR_004.PDF
'MAYOR_003.PDF

'INV_CUEN_



    Insertar = True
    
    'Si se agrupa y el numero presentacion es mayor que 0 para el diario y/o las cuentas anuales
    'Se a�aden dos valores mas, NIF e iVA
    'EN NIF tendremos la fecha final ejercicio si se cierra
        
        
    Cad = ",'" & Format(F1, "ddmmyyyy") & "','" & Format(F2, "ddmmyyyy") & "',"
    If Me.chkAgrupar.Value Then
        Cad = Cad & "'" & Format(DateAdd("d", -1, F1), "ddmmyyyy") & "')"
    Else
        Cad = Cad & "NULL)"
    End If
    
    'Le ponemos el numero en la BD tambien
    Cad = "," & numero & Cad
    
    
    If Me.chkAgrupar.Value = 0 Then
        Select Case opcion
        Case 0
            'Diario
            SQL = "DIARIO_" & Format(numero, "000") & ".PDF"
            Cad = "'DIARIO','Libro Diario'" & Cad
        Case 1
            SQL = "MAYOR_" & Format(numero, "000") & ".PDF"
            Cad = "'MAYOR','LIBRO MAYOR'" & Cad
        Case 2, 8
            SQL = "INVENTAR_" & Format(numero, "000") & ".PDF"
            If opcion = 2 Then
                Cad = "'INVENTAR','Inventario inicial'" & Cad
            Else
                Cad = "'INVENTAR','Inventario final Cierre'" & Cad
            End If
            
        Case 3
            SQL = "BAL_SUMS_" & Format(numero, "000") & ".PDF"
            Cad = "'" & "BALAN_" & numero & "','Balance sumas saldos'" & Cad
        Case 4
            SQL = "FAC_EMIT_" & Format(numero, "000") & ".PDF"
            Cad = "'" & "FACLI" & numero & "','Facturas emitidas'" & Cad
        Case 5
            SQL = "FAC_RECI_" & Format(numero, "000") & ".PDF"
            Cad = "'" & "FAPRO" & numero & "','Facturas recibidas'" & Cad
            
        Case 6
            'Perdidas y ganancias
            SQL = "PER_GAN_" & Format(numero, "000") & ".PDF"
            Cad = "'" & "PERGAN" & numero & "','P�rdidas y ganancias'" & Cad
            
        Case 7
            'Situacion
            SQL = "BALANCES_" & Format(numero, "000") & ".PDF"
            Cad = "'" & "BALANCES" & numero & "','Balance situacion'" & Cad
        End Select
        
    Else
        'AGRUPAMOS LOS LIBROS
        'Es decir solo el diario y el agrupado
        Select Case opcion
        Case 0
            SQL = "DIARIO_" & Format(numero, "000") & ".PDF"
            Cad = "'DIARIO','Libro Diario'" & Cad
        Case 100
            'Es el libro conjunto
            SQL = "INV_CUEN_" & Format(numero, "000") & ".PDF"
            Cad = "'TODO','Inventario y cuentas anuales'" & Cad
            
        Case Else
        
            'EL RESTO DE LIBROS
            'los copiamos en la tmp
            Insertar = False
        End Select
        
    End If
    lblIndicador.Caption = "Copiar archivo: " & SQL
    lblIndicador.Refresh
    
    'Insertamos en tmpfaclin
    If Insertar Then
        Contador = Contador + 1
        Cad = "INSERT INTO tmpfaclin (codusu, codigo, Numfac,Cliente, iva,Fecha, cta,nif) VALUES (" & vUsu.Codigo & "," & Contador & "," & Cad
        Conn.Execute Cad
        
        Cad = vPa
        If Not AnyadirNombresTxt Then Exit Function
        
        'A�adimos el nombre al fichero de Nombres.Txt
        SQL = vPa & "\" & SQL
    
    Else
        'Son los libros que agruparemos
        If opcion <> 3 Then
            SQL = App.Path & "\Temp\" & opcion & "1.pdf"
        Else
            SQL = App.Path & "\Temp\" & opcion & Format(numero, "000") & ".pdf"
        End If
    End If
    FileCopy App.Path & "\Docum.pdf", SQL
    
    
    If Not Insertar Then
        Select Case opcion
        Case 0
            'Diario
            SQL = "Libro diario"
        Case 1
            SQL = "Libro Mayor"
        Case 2
            SQL = "Inventario Inicial"
            
        Case 3
            SQL = "Balance sumas y saldos"
            
        Case 4
            SQL = "Facturas emitidas"
        Case 5
            SQL = "Facturas recibidas"
            
        Case 6
            'Perdidas y ganancias
            SQL = "P�rdidas y ganancias"
        Case 7
            'Situacion
            SQL = "Balance de situaci�n"
        Case 8
            SQL = "Inventario final CIERRE"
        End Select
        
        'Meto ahora la tapa
        CrearSeparadorTapa True, SQL, CInt(opcion)
        
    End If
    
    CopiarArchivoLega = True
Exit Function
eCopiarArchivoLega:
    MuestraError Err.Number, "Copiando archivo"
End Function



Private Function AnyadirNombresTxt() As Boolean
    
        On Error GoTo EAnyadirNombresTxt
        AnyadirNombresTxt = False
        Cad = Cad & "\NOMBRES.TXT"
        Cont = FreeFile
        Open Cad For Append As #Cont
        Print #Cont, SQL
        Close #Cont
        AnyadirNombresTxt = True
        Exit Function
EAnyadirNombresTxt:
    MuestraError Err.Number, "Anyadir Nombres.Txt"
End Function

'Generamos el archivo DATOS
Private Sub CreaDatosTXT()
Dim i As Integer
Dim F As Date
Dim TieneInventario As Boolean

    Cad = Cad & "\DATOS.TXT"
    Cont = FreeFile
    Open Cad For Output As #Cont
    
'ESTRUCTURA DEL FICHERO
'100VALENCIA
    Print #Cont, "100" & "VALENCIA"
    
'Fecha presentacion
    'Para la fecha presentacion
    F = CDate(Text3(0).Text)
    F = DateAdd("yyyy", 1, F)
    i = Year(F)
    If Month(F) > 3 Then i = i + 1
    F = CDate("31/03/" & i)
    Print #Cont, "101" & Format(F, "ddmmyyyy")
    
'102David
    Print #Cont, "102" & txtDatos(1).Text
'103Gandul
    Print #Cont, "103"
'104Castells
    Print #Cont, "104"
    
'10524348588Y
    Print #Cont, "105" & txtDatos(0).Text
'106Avd
    Print #Cont, "106" & txtDatos(2).Text
'107Valencia
    Print #Cont, "107" & txtDatos(4).Text
'10846016
    Print #Cont, "108" & txtDatos(3).Text
'401NO
    Print #Cont, "401" & "NO"
'10946
    Print #Cont, "109" & "46"   'PRovincia
'110fax
    Print #Cont, "110"
'111654649836
    Print #Cont, "111"

'2011
    Print #Cont, "201"
'2041
    Print #Cont, "204"
'2061
    Print #Cont, "206"
'205REGISTRO MERCANTIL
    Print #Cont, "205" & "REGISTRO MERCANTIL"

'AHora van Numerando los libros
' 00n   Libro
'  y dentro de cada libro
'       01: Desc
'       02: Numero
'       03: F INIC
'       04: F FIN
'       05: Fecha cierre
'       06: FIRMA
'00101Balances comprobaci�n(sumasaldo)
'001022
'0010301012002
'0010431122002
'0010531122001
'001067FBGMDQHTSRKSRA0U2JF2XRE3F
'00201Diario
'002021
'0020301012002
'0020431122002
'00205
'0020603YBX20BEV510TXS51K8RU2Z0Y
'00301Facturas emitidas
'003025
'0030301012002
'0030431122002
'00305
'0030627MC1X4UHFC2V5TFGH87NSPKJV
'00401Facturas recibidas
'004026
'0040301012002
'0040431122002
'00405
'0040627MC1X4UHFC2V5TFGH87NSPKJV
'00501Inventario
'005024
'0050301012002
'0050401012002
'00505
'005062JRFSATT2VUK2P67DH0DG0S1U5
'00601Mayor
'006023
'0060301012002
'0060431122002
'0060531122001
'006065KU5HZ3E73MXF1614EMF13JPAB
    
        Cad = "Select * from tmpfaclin where codusu = " & vUsu.Codigo & " ORDER BY Codigo"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        NumRegElim = 0
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            miRsAux.MoveNext
        Wend
        
        '501 Numero total de libros
        Print #Cont, "501" & NumRegElim
        miRsAux.MoveFirst
        NumRegElim = 0
        TieneInventario = False
        While Not miRsAux.EOF
                ' 00n   Libro
                '  y dentro de cada libro
                '       01: Desc
                '       02: Numero
                '       03: F INIC
                '       04: F FIN
                '       05: Fecha cierre. PONER si se agrupa, y el numero de presentacion mayor que 1. Esta en el campo NIF
                '       06: FIRMA
                
                NumRegElim = NumRegElim + 1
                Cad = Format(NumRegElim, "000")
                '1
                SQL = Cad & "01" & miRsAux.Fields(5)
                Print #Cont, SQL
                'N� Libro. Para BALAN, si hay mas de uno
                SQL = Mid(miRsAux.Fields(2), 1, 6)
                
                
                Select Case SQL
                Case "BALAN_"
                    i = InStr(1, miRsAux.Fields(2), "_")
                    SQL = Mid(miRsAux.Fields(2), i + 1)
                    i = Val(SQL)
                Case "INVENT"
                        If TieneInventario Then
                            i = 2
                        Else
                            i = 1
                            TieneInventario = True
                        End If
                
                Case "DIARIO", "TODO"
                    i = DBLet(miRsAux!iva, "N")
                
                
                Case Else
                    i = 1
                    
                End Select
                
                
                
                SQL = Cad & "02" & i
                Print #Cont, SQL
                '3
                SQL = Cad & "03" & miRsAux.Fields(3).Value
                Print #Cont, SQL
                '4
                SQL = Cad & "04" & miRsAux.Fields(4).Value
                Print #Cont, SQL
                '5
                
                SQL = Cad & "05" & DBLet(miRsAux!NIF, "T")
                Print #Cont, SQL



            'Siguiente libro
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        Close (Cont)
        Set miRsAux = Nothing
        
        
        ' Este trozo era lo que ahora pone el SELECT CASE
        'ANTES
'                If SQL = "BALAN_" Then
'                    I = InStr(1, miRsAux.Fields(2), "_")
'                    SQL = Mid(miRsAux.Fields(2), I + 1)
'                    I = Val(SQL)
'                Else
'                    If SQL = "INVEN" Then
'                        If TieneInventario Then
'                            I = 2
'                        Else
'                            I = 1
'                            TieneInventario = True
'                        End If
'                    Else
'                        If SQL = "DIARIO" Then
'                            If Me.chkAgrupar.Value Then I = DBLet(miRsAux!iva, "N")
'                        Else
'                            I = 1
'                        End If
'                    End If
'                End If
        
        
End Sub



Private Function CrearSeparadorTapa(Separador As Boolean, Titul As String, numero As Integer) As Boolean
    Cad = App.Path & "\Informes\"
    If Separador Then
        SQL = "Desc= """ & Titul & """|"
        Cad = Cad & "separador.rpt"
    Else
        SQL = DateAdd("d", -1, DateAdd("yyyy", 1, CDate(Text3(0).Text)))
        SQL = "Desc= """ & Text3(0).Text & " - " & SQL & """|"
        Cad = Cad & "tapa.rpt"
            
    End If
    With frmVisReportN
        .SoloImprimir = False
        .OtrosParametros = SQL
        .FormulaSeleccion = ""
        .NumeroParametros = 1
        .MostrarTree = False
        .Informe = Cad
        .ExportarPDF = True
        .Show vbModal
    End With
    SQL = App.Path & "\Temp\" & numero & "0.pdf"
    'Copiamos el archivo a la carpeta
    FileCopy App.Path & "\Docum.pdf", SQL
    Me.Refresh
    espera 0.5
    Me.lblIndicador.Caption = "Generando datos"
    Me.Refresh
    Me.Show
    espera 2
    
End Function


Private Function CopiarDocum() As Boolean
    
    CopiarDocum = False
    Cont = 0
    Cad = ""
    lblIndicador.Caption = "Generando fichero Docu2.pdf  "
    Me.Refresh
    Do
        Cad = Dir(App.Path & "\Docum2.pdf", vbArchive)
        If Cad = "" Then
            Cont = Cont + 1
            If Cont > 10 Then
                Cad = "Transcurrido 10 segundos no finaliza el proceso. �Salir?"
                If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then
                    Cad = ""
                Else
                    Exit Function
                End If
            Else
                espera 1
            End If
        End If
                
    Loop Until Cad <> ""
    espera 1

    'Si llega aqui haremos tres intentos de copiar el archivo
    On Error Resume Next
    Cont = 0
    Do
        lblIndicador.Caption = "Copia fichero Docu2.pdf  (" & Cont & ")"
        Me.Refresh

        FileCopy App.Path & "\Docum2.pdf", App.Path & "\Docum.pdf"
        If Err.Number <> 0 Then
            
            Err.Clear
            Cont = Cont + 1
            espera CInt((Cont * 2)) + 1
        Else
            Cont = 100
        End If
    Loop Until Cont > 3
    
    If Cont = 100 Then
        CopiarDocum = True
        lblIndicador.Caption = "Eliminando Docu2.pdf"
        Me.Refresh
        Kill App.Path & "\Docum2.pdf"
    End If
End Function

