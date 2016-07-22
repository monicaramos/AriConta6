VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESRemesasTP 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remesas"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16035
   Icon            =   "frmTESRemesasTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   16035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   30
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   15915
      Begin VB.Frame FrameFiltro 
         Height          =   705
         Left            =   12330
         TabIndex        =   70
         Top             =   180
         Width           =   2415
         Begin VB.ComboBox cboFiltro 
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
            ItemData        =   "frmTESRemesasTP.frx":000C
            Left            =   60
            List            =   "frmTESRemesasTP.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   210
            Width           =   2235
         End
      End
      Begin VB.Frame FrameBotonGnral2 
         Height          =   705
         Left            =   3930
         TabIndex        =   27
         Top             =   180
         Width           =   1515
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   210
            TabIndex        =   28
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Abono Remesa"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Devolución"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   21
         Top             =   180
         Width           =   3495
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   180
            TabIndex        =   1
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
         Height          =   7305
         Left            =   240
         TabIndex        =   22
         Top             =   990
         Width           =   15525
         _ExtentX        =   27384
         _ExtentY        =   12885
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
         Left            =   14670
         TabIndex        =   0
         Top             =   7860
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6930
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCreacionRemesa 
      BorderStyle     =   0  'None
      Height          =   9045
      Left            =   30
      TabIndex        =   19
      Top             =   -60
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
         Left            =   10890
         TabIndex        =   57
         Tag             =   "Importe|N|N|||reclama|importes|||"
         Top             =   8550
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   4515
         Left            =   150
         TabIndex        =   26
         Top             =   3840
         Width           =   15495
         Begin MSComctlLib.ListView lwCobros 
            Height          =   1725
            Left            =   0
            TabIndex        =   76
            Top             =   270
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   3043
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
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
               Text            =   "Código"
               Object.Width           =   1765
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Referencia Talón"
               Object.Width           =   5293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Banco"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "F. Recepcion"
               Object.Width           =   2734
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "F. Vto"
               Object.Width           =   2734
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Cliente"
               Object.Width           =   6677
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Importe"
               Object.Width           =   3089
            EndProperty
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2025
            Left            =   0
            TabIndex        =   77
            Top             =   2430
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   3572
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Serie"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Factura"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "F. Factura"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Vto"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Fec. Vto"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Cuenta"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Cliente"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Importe"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Num tal"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label Label3 
            Caption         =   "Vencimientos"
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
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   75
            Top             =   2100
            Width           =   2280
         End
         Begin VB.Label Label3 
            Caption         =   "Documentos"
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
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   2280
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   14850
            Picture         =   "frmTESRemesasTP.frx":0050
            ToolTipText     =   "Puntear al Debe"
            Top             =   30
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   14490
            Picture         =   "frmTESRemesasTP.frx":019A
            ToolTipText     =   "Quitar al Debe"
            Top             =   30
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   180
         TabIndex        =   24
         Top             =   8340
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
         TabIndex        =   17
         Top             =   8460
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
         TabIndex        =   18
         Top             =   8460
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
      Begin VB.Frame FrameCreaRem 
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
         TabIndex        =   29
         Top             =   60
         Width           =   15645
         Begin VB.ComboBox cmbRemesa 
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
            ItemData        =   "frmTESRemesasTP.frx":02E4
            Left            =   11970
            List            =   "frmTESRemesasTP.frx":02EB
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   810
            Width           =   1695
         End
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
            TabIndex        =   14
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
            TabIndex        =   55
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
            TabIndex        =   16
            Tag             =   "Descripción|T|N|||remesas|descripción|||"
            Top             =   3150
            Width           =   5025
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
            TabIndex        =   15
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
            TabIndex        =   4
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
            TabIndex        =   5
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   3
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
            TabIndex        =   2
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
            TabIndex        =   11
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
            TabIndex        =   10
            Tag             =   "imgConcepto"
            Top             =   1950
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de remesa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   11
            Left            =   11970
            TabIndex        =   73
            Top             =   540
            Width           =   1740
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
            TabIndex        =   56
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
            Top             =   510
            Width           =   960
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Recepción"
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
            Index           =   8
            Left            =   2700
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
            Top             =   2010
            Width           =   690
         End
      End
      Begin VB.Frame FrameModRem 
         Caption         =   "Datos Remesa"
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
         TabIndex        =   59
         Top             =   60
         Width           =   15645
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
            Index           =   5
            Left            =   5370
            TabIndex        =   61
            Tag             =   "Fecha Reclamación|F|N|||reclama|fecreclama|dd/mm/yyyy||"
            Text            =   "99/99/9999"
            Top             =   2130
            Width           =   1245
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
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   62
            Tag             =   "Descripción|T|N|||remesas|descripción|||"
            Top             =   2130
            Width           =   5025
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
            Index           =   3
            Left            =   1740
            TabIndex        =   63
            Text            =   "Text2"
            Top             =   2130
            Width           =   3525
         End
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
            Index           =   3
            Left            =   360
            TabIndex        =   60
            Text            =   "Text2"
            Top             =   2130
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Remesa"
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
            Height          =   375
            Index           =   12
            Left            =   360
            TabIndex        =   68
            Top             =   750
            Width           =   8940
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
            Index           =   0
            Left            =   2580
            TabIndex        =   67
            Top             =   3990
            Width           =   4095
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
            Index           =   4
            Left            =   5370
            TabIndex        =   66
            Top             =   1860
            Width           =   795
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   6330
            Top             =   1860
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
            Index           =   3
            Left            =   6690
            TabIndex        =   65
            Top             =   1860
            Width           =   1245
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   3
            Left            =   2340
            Top             =   1860
            Width           =   240
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
            Index           =   2
            Left            =   360
            TabIndex        =   64
            Top             =   1860
            Width           =   1845
         End
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   0
         Left            =   2130
         TabIndex        =   69
         Top             =   8550
         Width           =   6960
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
         Left            =   9210
         TabIndex        =   58
         Top             =   8610
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmTESRemesasTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)



Private Const SaltoLinea = """ + chr(13) + """

Private Const IdPrograma = 612


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
Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1

Private frmMens3 As frmMensajes
Private frmMens2 As frmMensajes
Attribute frmMens2.VB_VarHelpID = -1
Private frmMens As frmMensajes

Dim SQL As String
Dim RC As String
Dim RS As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim CONT As Long
Dim I As Integer
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

Dim Indice As Integer
Dim Codigo As Long

Dim SubTipo As Integer

Dim ModoInsertar As Boolean

Dim IndCodigo As Integer

Dim SelTalones As Boolean
Dim SelPagares As Boolean



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


Private Sub cboFiltro_Click()
    CargarSqlFiltro
    CargaList
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
Dim I As Integer

    If Index = 0 Then
    
        If ModoInsertar Then
            cmdAceptar(0).Caption = "&Aceptar"
            ModoInsertar = False
        End If
    
        Frame1.Visible = True
        Frame1.Enabled = True
        
        FrameCreacionRemesa.Visible = False
        FrameCreacionRemesa.Enabled = False
        If I >= 0 Then lw1.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
    Select Case Index
        Case 0
            Select Case Modo
                Case 3  ' insertar
                    If Not DatosOK(0) Then Exit Sub
                
                    If Not ModoInsertar Then
                        ModoInsertar = True
                        cmdAceptar(0).Caption = "C&onfirmar"
                        If SubTipo <> vbTipoPagoRemesa Then
                            NuevaRemTalPag
                        Else
                            NuevaRem
                        End If
                    Else
                        If EfectuarRemesa(0) Then
                            MsgBox "Remesa generada correctamente.", vbExclamation
                            cmdCancelar_Click (0)
                            CargaList
                        End If
                    End If
                    
                    Screen.MousePointer = vbDefault
                    
                    
                Case 4  ' modificar
                    If Not DatosOK(1) Then Exit Sub
                    
                    If Not ModoInsertar Then
                        ModoInsertar = True
                        
                        cmdAceptar(0).Caption = "C&onfirmar"
                    Else
                        If EfectuarRemesa(1) Then
                            'Refrescamos los datos en el lw de remesas
                            'MsgBox "Remesa modificada correctamente.", vbExclamation
                            cmdCancelar_Click (0)
                        End If
                        
                    End If
            End Select
    End Select
End Sub


Private Function DatosOK(Opcion As Integer) As Boolean
Dim B As Boolean
Dim CtaPuente As Boolean
Dim YaRemesado As Currency
Dim Limite As Currency

    DatosOK = False

    If Opcion = 0 Then ' insertar remesa
        If txtCuentas(2).Text = "" Then
            MsgBox "Indique la cuenta bancaria", vbExclamation
            Exit Function
        Else
            SQL = "select count(*) from bancos where codmacta = " & DBSet(txtCuentas(2).Text, "T") & " and not sufijoem is null and sufijoem <> ''"
            If TotalRegistros(SQL) = 0 Then
                MsgBox "El banco no tiene Sufijo OEM. Reintroduzca.", vbExclamation
                PonleFoco txtCuentas(2)
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
        
        
        If Me.cmbRemesa.ListIndex = 0 Then
            CtaPuente = vParamT.PagaresCtaPuente
        Else
            CtaPuente = vParamT.TalonesCtaPuente
        End If
        
        If txtRemesa.Text = "" Then
            MsgBox "El campo descripcion debe tener valor", vbExclamation
            Exit Function
        End If

        If ModoInsertar = 1 Then
            
        
            'Crear remesa talon pagare
            Importe = 0
            For NumRegElim = 1 To lwCobros.ListItems.Count
                If lwCobros.ListItems(NumRegElim).Checked Then
                    'Este documento. Vemos el importe del documento
                    'Antes septiembre 2009
                    'SQL = DevuelveDesdeBD("sum(importe)", "slirecepdoc", "id", ListView2.ListItems(NumRegElim).Text)
                    'If SQL = "" Then SQL = "0"   'No deberia pasar
                    SQL = lwCobros.ListItems(NumRegElim).SubItems(6)
                    Importe = Importe + CCur(SQL)
                End If
            Next
    
            If Importe = 0 Then
                MsgBox "Seleccione algun talón/pagaré", vbExclamation
                Exit Function
            End If
    
            'La fecha y las cuentas bloqueadas ya las hemos comprobado en la fase anterior
            'Ahora el limite del banco
            If cmbRemesa.ListIndex = 1 Then
                SQL = "talonriesgo"
                NumRegElim = 3 '   para la select de abajo
            Else
                NumRegElim = 2
                SQL = "pagareriesgo" 'para la select de abajo
            End If
            SQL = DevuelveDesdeBD(SQL, "bancos", "codmacta", Trim(Mid(Text1(0).Text, 1, 10)), "T")
            If SQL <> "" Then
                Limite = CCur(SQL)
            Else
                Limite = -1
            End If
            
            'Tenemos que ver todos los vencimientos que sean de tipo de pago talon o pagare, que la cta de pago sea
            'la del banco en question y ver cuanto llevamos remesado
            SQL = "select sum(impcobro) FROM cobros,formapago WHERE cobros.codforpa = formapago.codforpa AND "
            SQL = SQL & "siturem>'B' AND siturem < 'Z'"
    
            SQL = SQL & " and ctabanc1='" & Trim(Mid(Text1(0).Text, 1, 10)) & "' AND tiporem = " & NumRegElim
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            YaRemesado = 0
            If Not miRsAux.EOF Then
                'Le sumo lo que llevamos en esta remesa (los k estan check) a los vtos ya remesados Y nO eleminidados
                YaRemesado = DBLet(miRsAux.Fields(0), "N")
            End If
            miRsAux.Close
            Set miRsAux = Nothing
            
            If Limite >= 0 Then
                If Limite - (Importe + YaRemesado) < 0 Then
                    'Supera el riesgo
                    SQL = "Esta superando el riesgo permitido por el banco." & vbCrLf
                    SQL = SQL & "Riesgo concedido: " & Format(Limite, FormatoImporte) & vbCrLf
                    SQL = SQL & "Remesa: " & Format(Importe, FormatoImporte) & vbCrLf
                    SQL = SQL & "Ya remesado: " & Format(YaRemesado, FormatoImporte) & vbCrLf
                    
                    SQL = SQL & "¿Continuar?"
                    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
                End If
            End If
            
            Set miRsAux = New ADODB.Recordset
            
            'UNa ultima comprobacion. Vamos a ver si un mismo vencimiento esta en dos docuemntos
            'distintos, o si alguno de los vencimientos pertecence a una remesa que aun no ha sido
            'borrada
            If Not ComprobarEfectosCobradosParcialmente Then
                Set miRsAux = New ADODB.Recordset
                Exit Function
            End If
        End If
    Else
        If txtCuentas(3).Text = "" Then
            MsgBox "Indique la cuenta bancaria", vbExclamation
            Exit Function
        End If
    
        'Fecha remesa tiene k tener valor
        If txtFecha(5).Text = "" Then
            MsgBox "Fecha de remesa debe tener valor", vbExclamation
            PonFoco txtFecha(5)
            Exit Function
        Else
            If Year(CDate(txtFecha(5).Text)) <> lw1.SelectedItem.SubItems(1) Then
                MsgBox "La fecha de remesa ha de ser del mismo año. Revise.", vbExclamation
                PonFoco txtFecha(5)
                Exit Function
            End If
        End If
        
        'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
        If FechaCorrecta2(CDate(txtFecha(5).Text), True) > 1 Then
            PonFoco txtFecha(5)
            Exit Function
        End If
    End If
    
    DatosOK = True

End Function

Private Sub Insertar()
Dim NumF As Long
Dim B As Boolean

    On Error GoTo eInsertar
    
    Conn.BeginTrans
    
eInsertar:
    If Err.Number = 0 And B Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
End Sub

Private Function InsertarLineas() As Boolean
Dim RS As ADODB.Recordset
Dim CadValues As String
Dim CadInsert As String

    On Error GoTo eInsertarLineas

    InsertarLineas = False

    InsertarLineas = True
    Exit Function
    
eInsertarLineas:
    MuestraError Err.Number, "Insertar Lineas", Err.Description
End Function

Private Sub cmdVtoDestino(Index As Integer)
    
    If Index = 0 Then
        TotalReg = 0
        If Not Me.lwCobros.SelectedItem Is Nothing Then TotalReg = Me.lwCobros.SelectedItem.Index
    
    
        For I = 1 To Me.lwCobros.ListItems.Count
            If Me.lwCobros.ListItems(I).Bold Then
                Me.lwCobros.ListItems(I).Bold = False
                Me.lwCobros.ListItems(I).ForeColor = vbBlack
                For CONT = 1 To Me.lwCobros.ColumnHeaders.Count - 1
                    Me.lwCobros.ListItems(I).ListSubItems(CONT).ForeColor = vbBlack
                    Me.lwCobros.ListItems(I).ListSubItems(CONT).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalReg > 0 Then
            I = TotalReg
            Me.lwCobros.ListItems(I).Bold = True
            Me.lwCobros.ListItems(I).ForeColor = vbRed
            For CONT = 1 To Me.lwCobros.ColumnHeaders.Count - 1
                Me.lwCobros.ListItems(I).ListSubItems(CONT).ForeColor = vbRed
                Me.lwCobros.ListItems(I).ListSubItems(CONT).Bold = True
            Next
        End If
        lwCobros.Refresh
        
        PonerFocoLw Me.lwCobros

    Else
    

    End If
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Not Frame1.Visible Then
            If CadenaDesdeOtroForm <> "" Then
            Else
'                PonFoco Text1(2)
            End If
            CadenaDesdeOtroForm = ""
        End If
        
        cboFiltro.ListIndex = 0
        
        CargaList
        PonerFocoBtn cmdCancelar(1)
        If lw1.ListItems.Count > 0 Then Set lw1.SelectedItem = Nothing

    End If
    Screen.MousePointer = vbDefault
End Sub
    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmPpal.Icon
    
    For I = 0 To 1
        Me.imgSerie(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgCuentas(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    Me.imgCuentas(2).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(3).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    
    For I = 0 To 5
        Me.imgFec(I).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next I
    
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
        .Buttons(1).Image = 37
        .Buttons(2).Image = 45
    End With
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    Caption = Caption & "       PAGARES y TALONES"
    
    CargaFiltros
    
    cmbRemesa.Clear
    cmbRemesa.AddItem "Pagarés"
    cmbRemesa.AddItem "Talones"
    
    cmbRemesa.ListIndex = 0
    
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
    
    
    Orden = True
    CampoOrden = "remesas.fecremesa"
    
    If Tipo = 1 Then
        SubTipo = vbTipoPagoRemesa
    Else
        SubTipo = vbTalon
    End If
    
End Sub


Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    
    Select Case Me.cboFiltro.ListIndex
        Case 0 ' sin filtro
            SelTalones = True
            SelPagares = True
        
        Case 1 ' pagares
            SelTalones = False
            SelPagares = True
        
        Case 2 ' talones
            SelTalones = True
            SelPagares = False
        
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCuentas(2).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNCuentas(2).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
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
Dim I As Integer
    For I = 1 To Me.lwCobros.ListItems.Count
        Set IT = lwCobros.ListItems(I)
        lwCobros.ListItems(I).Checked = (Index = 1)
        lwCobros_ItemCheck (IT)
        Set IT = Nothing
    Next I
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'FECHA FACTURA
    Indice = Index
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtFecha(Indice).Text <> "" Then frmF.Fecha = CDate(txtFecha(Indice).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco txtFecha(Indice)

End Sub

Private Sub imgCuentas_Click(Index As Integer)
    
    If Index = 2 Then
            Set frmBan = New frmBasico2
            AyudaBanco frmBan
            Set frmBan = Nothing
    
    
    Else
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
    End If
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2, 3, 4, 5
        Indice = Index
    
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

Private Sub lw1_DblClick()
    'detalle de facturas
    Set frmMens = New frmMensajes
    
    frmMens.Opcion = 51
    frmMens.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|"
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    
End Sub

Private Sub lw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'FALTA###  ¿Porque esta asi?
'    PonerModoUsuarioGnral 2, "ariconta"
End Sub

Private Sub lwCobros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim C As Currency
Dim Cobro As Boolean

    Cobro = True
'    C = Item.Tag
    
    Importe = 0
    For I = 1 To lwCobros.ListItems.Count
        If lwCobros.ListItems(I).Checked Then Importe = Importe + lwCobros.ListItems(I).SubItems(6)
    Next I
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
            If Not lw1.SelectedItem Is Nothing Then
                frmTESRemesasTPList.numero = lw1.SelectedItem.Text
                frmTESRemesasTPList.Anyo = lw1.SelectedItem.SubItems(1)
            Else
                frmTESRemesasTPList.numero = ""
                frmTESRemesasTPList.Anyo = ""
            End If
            frmTESRemesasTPList.Show vbModal

    End Select
End Sub

Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    
    SepuedeBorrar = False

    If lw1.SelectedItem.SubItems(8) = "Q" Then
        MsgBox "No se pueden modificar ni eliminar remesas en situación abonada.", vbExclamation
        Exit Function
    End If
    
    SepuedeBorrar = True

End Function


Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    
    
    If Me.lw1.SelectedItem = "" Then Exit Sub
        
    If Not SepuedeBorrar Then Exit Sub
        
        
    '*************** canviar els noms i el DELETE **********************************
    
    SQL = lw1.SelectedItem.SubItems(10)
    SQL = vbCrLf & "Tipo :  " & SQL
    SQL = "¿Seguro que desea eliminar la Remesa?" & SQL
    SQL = SQL & vbCrLf & " Código: " & lw1.SelectedItem.Text
    SQL = SQL & vbCrLf & " Fecha: " & lw1.SelectedItem.SubItems(2)
    SQL = SQL & vbCrLf & " Banco: " & lw1.SelectedItem.SubItems(5)
    SQL = SQL & vbCrLf & " Importe: " & lw1.SelectedItem.SubItems(7)
    
    
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        If BloqueoManual(True, "Remesas", "Remesas") Then
            'Hay que eliminar
            
            If Tipo = 1 Then
            
            
                SQL = "Delete from remesas where codigo=" & lw1.SelectedItem.Text
                SQL = SQL & " AND anyo =" & lw1.SelectedItem.SubItems(1)
                SQL = SQL & " AND tiporem =" & lw1.SelectedItem.SubItems(9)
                Conn.Execute SQL
            
            
            
                'Agosto2013  Ponemos a null la cuenta real de cobroctabanc2
                'Pongo A NULL todos los recibos con esos valores
                SQL = "UPDATE cobros set codrem=NULL,anyorem=NULL,siturem=NULL,tiporem=NULL"
                SQL = SQL & ",fecultco=NULL,impcobro=NULL,ctabanc2=NULL"
                SQL = SQL & " where codrem=" & lw1.SelectedItem.Text
                SQL = SQL & " AND anyorem =" & lw1.SelectedItem.SubItems(1)
                SQL = SQL & " AND tiporem =" & lw1.SelectedItem.SubItems(9)
                Conn.Execute SQL
            
            Else
                If BorrarRemesaEnCancelacionTalonesPagares Then CargaList
            End If
            'adodc1.Recordset.Cancel
            BloqueoManual False, "Remesas", ""
        
        Else
            MsgBox "Proceso bloqueado por otro usuario", vbExclamation
        End If

    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Function BorrarRemesaEnCancelacionTalonesPagares() As Boolean
Dim C As String

    On Error GoTo EBorrarRemesaEnCancelacionTalonesPagares

    'En cancelacion si que dejo eliminar, ya que lo que se hace realmente es:
    '1.- QUitar la remesa de los cobros       'Estos dos puntos los hace en la otra
    '2.- Quitar la remesa de la tabla remesas
    '3.- poner en scarecepdoc LlevadoBanco=0
        
    BorrarRemesaEnCancelacionTalonesPagares = False

    'Veamos que scarecep son
    Set miRsAux = New ADODB.Recordset
    C = "select codigo from talones_facturas where (numserie,numfactu,fecfactu,numorden) IN ("
    C = C & "SELECT numserie,numfactu,fecfactu,numorden FROM cobros WHERE "
    C = C & " codrem=" & lw1.SelectedItem.Text & " AND anyorem = " & lw1.SelectedItem.SubItems(1) & ")"
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        C = "UPDATE talones set LlevadoBanco = 0 WHERE codigo = " & miRsAux!Codigo
        Conn.Execute C
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Ponemos los vencimientos sin remesa
    C = "UPDATE cobros SET codrem=NULL, anyorem=NULL,siturem=NULL where"
    C = C & " codrem=" & lw1.SelectedItem.Text & " AND anyorem = " & lw1.SelectedItem.SubItems(1)
    Conn.Execute C
    
    'Borramos la remesa
    C = "DELETE from remesas WHERE "
    C = C & " Codigo=" & lw1.SelectedItem.Text & " AND Anyo = " & lw1.SelectedItem.SubItems(1)
    Conn.Execute C
    
    BorrarRemesaEnCancelacionTalonesPagares = True
    Exit Function
EBorrarRemesaEnCancelacionTalonesPagares:
    MsgBox "Error grave. Consulte soporte técnico", vbExclamation
End Function




Private Function ModificarCobros() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim FecUltCob As String
Dim Importe As Currency
Dim NumLinea As Integer


    ModificarCobros = False
    
    Conn.BeginTrans

    SQL = "select * from cobros where codrem = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    SQL = SQL & " and anyorem = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        
        ' antes lo sumaba de los cobros_realizados
        ' ahora lo dejo todo a nulo
    
    
        FecUltCob = ""
        Importe = 0
    
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
        SQL = SQL & ", situacion = 0 "
        SQL = SQL & " where numserie = " & DBSet(RS!NUmSerie, "T") & " and "
        SQL = SQL & " numfactu = " & DBSet(RS!NumFactu, "N") & " and fecfactu = " & DBSet(RS!FecFactu, "F") & " and "
        SQL = SQL & " numorden = " & DBSet(RS!numorden, "N")
                    
        Conn.Execute SQL
    
        RS.MoveNext
    Wend

    SQL = "delete from remesas where codigo = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    SQL = SQL & " and anyo = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Conn.Execute SQL

    Set RS = Nothing
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

    
    ModoInsertar = False
    
    LimpiarCampos
    Modo = 3
    PonerModo Modo

    txtFecha(4).Text = Format(Now, "dd/mm/yyyy")

    txtCuentas(2).Text = BancoPropio
    If txtCuentas(2).Text <> "" Then
        txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "bancos", "descripcion", "codmacta", txtCuentas(2), "T")
        If txtNCuentas(2).Text = "" Then txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtCuentas(2).Text, "T")
    End If

    PonleFoco txtFecha(2)
    
'    Label2.Caption = ""

    Me.Label3(8).Caption = "Fecha recepcion"
    Label1(1).Caption = "Banco remesar"



End Sub

Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    
    Me.lwCobros.ListItems.Clear
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub BotonModificar()
Dim SQL As String
    
 
    If lw1.SelectedItem Is Nothing Then Exit Sub
    If lw1.SelectedItem = 0 Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    ModoInsertar = True
    
    LimpiarCampos
    
    Modo = 4
    PonerModo Modo

    txtFecha(5).Text = Format(lw1.SelectedItem.SubItems(2), "dd/mm/yyyy")
    txtCuentas(3).Text = lw1.SelectedItem.SubItems(4)
    txtNCuentas(3).Text = lw1.SelectedItem.SubItems(5)
    Text2.Text = lw1.SelectedItem.SubItems(6)

    Label3(12).Caption = "Remesa: " & lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1) & ""

    If SubTipo = vbTipoPagoRemesa Then
        Me.Label3(8).Caption = "Fecha factura"
        Label1(1).Caption = "Banco"
    Else
        Me.Label3(8).Caption = "Fecha recepción"
        Label1(1).Caption = "Banco remesar"
    End If
    
    SQL = "from cobros, formapago where codrem = " & DBSet(lw1.SelectedItem.Text, "N") & " and anyorem = " & lw1.SelectedItem.SubItems(1)
    SQL = SQL & " and cobros.codforpa = formapago.codforpa"
    
    PonerVtosRemesa SQL, False

    PonleFoco txtCuentas(3)

End Sub




Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    PonerIndicador lblIndicador, Modo
    
    If Modo = 3 Or Modo = 4 Then
        Frame1.Visible = False
        Frame1.Enabled = False
    
        Me.FrameCreacionRemesa.Visible = True
        Me.FrameCreacionRemesa.Enabled = True
    End If
    
    If Modo = 3 Then
        Me.FrameCreaRem.Visible = True
        Me.FrameCreaRem.Enabled = True
        
        Me.FrameModRem.Visible = False
        Me.FrameModRem.Enabled = False
    Else
        Me.FrameCreaRem.Visible = False
        Me.FrameCreaRem.Enabled = False
        
        Me.FrameModRem.Visible = True
        Me.FrameModRem.Enabled = True
    End If
    
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar2 Button.Index
End Sub

Private Sub HacerToolBar2(Boton As Integer)

    Select Case Boton
        Case 1 ' CONTABILIZACION REMESA
            
            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
            
            
            HaHabidoCambios = False
        
            SQL = "No se puede contabilizar una "
            CadenaDesdeOtroForm = ""
            If lw1.SelectedItem.SubItems(8) = "A" Then CadenaDesdeOtroForm = SQL & "Remesa abierta. Sin llevar al banco."
            'Ya contabilizada
            If lw1.SelectedItem.SubItems(8) = "Q" Then CadenaDesdeOtroForm = SQL & "Remesa abonada."
            If CadenaDesdeOtroForm <> "" Then
                MsgBox CadenaDesdeOtroForm, vbExclamation
                CadenaDesdeOtroForm = ""
                Exit Sub
            End If
            CadenaDesdeOtroForm = ""
            
            frmTESRemesasTPCont.Opcion = 8
            frmTESRemesasTPCont.NumeroDocumento = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(4) & "|" & lw1.SelectedItem.SubItems(5) & "|" & lw1.SelectedItem.SubItems(7) & "|"
            frmTESRemesasTPCont.Show vbModal
         
            'Hay que poner en el formualrio de arriba valor a cadenadesdeotroform si ha modificado
            If HaHabidoCambios Then CargaList
         
        Case 2 ' DEVOLUCION DE REMESA
            HaHabidoCambios = False
             
            'FALTA####
            'Moniiiii, aqui mismo ya sabesmos en la situacion que esta la remesa. Si no es A o B
            '   ni seguimos
            If Not lw1.SelectedItem Is Nothing Then
                If Not Asc(lw1.SelectedItem.SubItems(8)) > Asc("B") Then
                    MsgBox "Remesa no se puede realizar la devolucion", vbExclamation
                    Exit Sub
                End If
            End If
                        
            frmTESRemesasTPDev.Opcion = 9
            frmTESRemesasTPDev.SubTipo = 1
            If Not lw1.SelectedItem Is Nothing Then
                frmTESRemesasTPDev.NumeroDocumento = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(4) & "|" & lw1.SelectedItem.SubItems(5) & "|" & lw1.SelectedItem.SubItems(7) & "|"
            Else
                frmTESRemesasTPDev.NumeroDocumento = ""
            End If
            frmTESRemesasTPDev.Show vbModal
         
            'Hay que poner en el formualrio de arriba valor a cadenadesdeotroform si ha modificado
            If HaHabidoCambios Then CargaList
         
         
    End Select
End Sub

Private Sub BorrarRemesaVtos()
Dim SQL As String

    NumRegElim = 0
    SQL = "Select count(*) from cobros where codrem=" & lw1.SelectedItem.Text
    SQL = SQL & " AND anyorem =" & lw1.SelectedItem.SubItems(1)
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    Set miRsAux = Nothing
    
    SQL = "Va a borrar la remesa y los vencimientos para: "
    SQL = SQL & vbCrLf & " --------------------------------------------------------------------"
    SQL = SQL & vbCrLf & "Código: " & lw1.SelectedItem.Text
    SQL = SQL & vbCrLf & "Año: " & lw1.SelectedItem.SubItems(1)
    SQL = SQL & vbCrLf & "Banco: " & lw1.SelectedItem.SubItems(4) & " " & lw1.SelectedItem.SubItems(5)
    SQL = SQL & vbCrLf & "Situación: " & lw1.SelectedItem.SubItems(3)
    SQL = SQL & vbCrLf & "Importe: " & Format(lw1.SelectedItem.SubItems(7), FormatoImporte)
    SQL = SQL & vbCrLf & "Vencimientos: " & NumRegElim
    SQL = SQL & vbCrLf & vbCrLf & "                         ¿Continuar?"
    NumRegElim = 0
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    SQL = "El proceso es irreversible"
    SQL = SQL & vbCrLf & "Desea continuar?"
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    Screen.MousePointer = vbHourglass
    If HacerEliminacionRemesaVtos Then
        'Cargar datos
        CargaList
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Function HacerEliminacionRemesaVtos() As Boolean

    On Error GoTo EHacerEliminacionRemesaVtos

    HacerEliminacionRemesaVtos = False

    'Eliminamos los vencimientos asociados
    Conn.Execute "DELETE FROM cobros where codrem=" & lw1.SelectedItem.Text & " AND anyorem =" & lw1.SelectedItem.SubItems(1)
    
    'Eliminamos la remesa
    Conn.Execute "DELETE FROM remesas where codigo=" & lw1.SelectedItem.Text & " AND anyo =" & lw1.SelectedItem.SubItems(1)
    
    HacerEliminacionRemesaVtos = True
    Exit Function
EHacerEliminacionRemesaVtos:
    MuestraError Err.Number, "Function: HacerEliminacionRemesaVtos"
End Function


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
    Me.Refresh
End Sub


Private Sub CargarDatos(vSQL As String, Modificar As Boolean)
    Dim IT
    
    lwCobros.ListItems.Clear
    
    SQL = "Select * from talones,cuentas where talones.codmacta = cuentas.codmacta AND " & vSQL
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = miRsAux!Codigo
        IT.SubItems(1) = miRsAux!numeroref
        IT.SubItems(2) = DBLet(miRsAux!Banco, "T") & " "
        IT.SubItems(3) = Format(miRsAux!fecharec, "dd/mm/yyyy")
        IT.SubItems(4) = Format(miRsAux!fechavto, "dd/mm/yyyy")
        IT.SubItems(5) = miRsAux!Nommacta
        IT.ListSubItems(5).ToolTipText = miRsAux!codmacta
        IT.SubItems(6) = Format(miRsAux!Importe, FormatoImporte)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If lwCobros.ListItems.Count > 0 Then
        Set lwCobros.SelectedItem = lwCobros.ListItems(1)
        CargaDatosLineas
    End If

End Sub


Private Sub CargaDatosLineas()
Dim IT As ListItem
Dim CodRem As Integer
Dim AnyoRem As Integer

    On Error GoTo EC
    
    If txtRemesa.Tag <> "" Then
        CodRem = txtRemesa.Tag
        AnyoRem = Year(CDate(txtFecha(4).Text))
    Else
        CodRem = 0
    End If
    
    ListView1.ListItems.Clear
    SQL = "Select cobros.numserie,cobros.numfactu,cobros.fecfactu,cobros.fecvenci, cobros.numorden,impvenci ,gastos ,impcobro,talones.numeroref reftalonpag,codrem,anyorem  "
    SQL = SQL & " FROM (talones inner join talones_facturas on talones.codigo = talones_facturas.codigo) left join cobros on cobros.numserie=talones_facturas.numserie AND talones_facturas.numfactu=cobros.numfactu and"
    SQL = SQL & " cobros.fecfactu=talones_facturas.fecfactu AND cobros.numorden=talones_facturas.numorden"
    SQL = SQL & " WHERE talones.codigo= " & lwCobros.SelectedItem.Text
        
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        If IsNull(miRsAux!NUmSerie) Then
            'ERROR GRAVE. Hay un vto en las lineas del docuemnto que NO esta en
            IT.ForeColor = vbRed
            IT.Bold = True
            IT.Text = "ERR"
            For NumRegElim = 1 To ListView1.ColumnHeaders.Count - 1
                IT.SubItems(NumRegElim) = "ERROR"
                IT.ListSubItems(NumRegElim).ForeColor = vbRed
                
                IT.ListSubItems(NumRegElim).Bold = True
            Next NumRegElim
        Else
             IT.Text = Mid(DBLet(miRsAux!NUmSerie, "T") & "   ", 1, 3)
             IT.SubItems(1) = Format(miRsAux!NumFactu, "000000")
             IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
             IT.SubItems(3) = miRsAux!numorden
             IT.SubItems(4) = Format(miRsAux!FecVenci, "dd/mm/yyyy")
            ' IT.SubItems(5) = miRsAux!codmacta
            ' IT.SubItems(6) = miRsAux!Nommacta
             'Lo debe cojer de impcobro
             IT.SubItems(7) = Format(miRsAux!impcobro, FormatoImporte)
             
             IT.SubItems(8) = DBLet(miRsAux!reftalonpag, "T")
             
             If CodRem > 0 Then
                 If Not IsNull(miRsAux!CodRem) Then
                     If Val(miRsAux!CodRem) = CodRem And Val(miRsAux!AnyoRem) = AnyoRem Then
                         'Voy a pintar de colorines el vto
                         IT.ForeColor = vbRed
                         For NumRegElim = 1 To IT.ListSubItems.Count
                             IT.ListSubItems(NumRegElim).ForeColor = vbRed
                         Next NumRegElim
                         IT.Checked = True
                     End If
                 End If
             End If
         End If 'de null numserie
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set ListView1.SelectedItem = Nothing
   
    Exit Sub
EC:
    MuestraError Err.Number, "Carga datos"
End Sub



Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    SQL = ""
    For I = 1 To Me.lwCobros.ListItems.Count
        If Me.lwCobros.ListItems(I).Checked Then
        
            Insertar = True
            If Me.lwCobros.ListItems(I).Bold Then
                RegistroDestino = I
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                SQL = SQL & ", ('" & lwCobros.ListItems(I).Text & "'," & lwCobros.ListItems(I).SubItems(1)
                SQL = SQL & ",'" & Format(lwCobros.ListItems(I).SubItems(2), FormatoFecha) & "'," & lwCobros.ListItems(I).SubItems(3) & ")"
            End If
            
        End If
    Next
    SQL = Mid(SQL, 2)
            
End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim RS As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(RS!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = DBLet(RS!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(RS!Imprimir, "N")
    
        Toolbar2.Buttons(1).Enabled = DBLet(RS!especial, "N") And Not (lw1.SelectedItem Is Nothing)
        Toolbar2.Buttons(2).Enabled = DBLet(RS!especial, "N") And Not (lw1.SelectedItem Is Nothing)
        Toolbar2.Buttons(3).Enabled = DBLet(RS!especial, "N") 'And Not (lw1.SelectedItem Is Nothing)
    End If
    
    RS.Close
    Set RS = Nothing
    
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
    
    If CampoOrden = "" Then CampoOrden = "remesas.anyo, remesas.codigo " 'remesas.fecremesa"
    Cad = Cad & " ORDER BY " & CampoOrden ' remesas.anyo desc,
    If Orden Then Cad = Cad & " DESC"
    
    lw1.ColumnHeaders.Clear
    
    lw1.ColumnHeaders.Add , , "Código", 950
    lw1.ColumnHeaders.Add , , "Año", 700
    lw1.ColumnHeaders.Add , , "Fecha", 1350
    lw1.ColumnHeaders.Add , , "Situación", 1640
    lw1.ColumnHeaders.Add , , "Cuenta", 1440
    lw1.ColumnHeaders.Add , , "Nombre", 2940
    lw1.ColumnHeaders.Add , , "Descripción", 2940
    lw1.ColumnHeaders.Add , , "Importe", 1940, 1
    lw1.ColumnHeaders.Add , , "S", 0, 1
    lw1.ColumnHeaders.Add , , "T", 0, 1
    lw1.ColumnHeaders.Add , , "Tipo", 1300
    
    
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
        IT.SubItems(9) = miRsAux!Tiporem
        
        If miRsAux!Tiporem = 2 Then
            IT.SubItems(10) = "PAGARE"
        Else
            IT.SubItems(10) = "TALON"
        End If
        
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
        'TALON Y PAGARE
        Select Case cboFiltro.ListIndex
            Case 0
                 SelTalones = True
                 SelPagares = True
            Case 1
                 SelTalones = True
                 SelPagares = False
            Case 2
                 SelTalones = False
                 SelPagares = True
        End Select
        C = RemesaSeleccionTipoRemesa(False, SelTalones, SelPagares)
    End If
    
    If C <> "" Then C = " AND " & C
    PonerOrdenFiltro = C
End Function



Private Sub NuevaRem()

Dim ForPa As String
Dim Cad As String
Dim Impor As Currency
Dim colCtas As Collection
Dim SQL2 As String

    If SubTipo = vbTipoPagoRemesa Then
        SQL = " formapago.tipforpa = " & vbTipoPagoRemesa
    Else
    
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
    
    SQL = SQL & " and situacion = 0 "
     
    ' si hay cobros con impcobro <> 0 damos aviso y no los incluimos
    If SubTipo = vbTipoPagoRemesa Then
    
        CadenaDesdeOtroForm = ""
    
        SQL2 = SQL & " and not cobros.impcobro is null and cobros.impcobro <> 0 and cobros.codmacta=cuentas.codmacta AND (siturem is null) AND cobros.codforpa = formapago.codforpa "
        
        SQL2 = "select cobros.* FROM cobros,cuentas,formapago  WHERE " & SQL2
        
        If TotalRegistrosConsulta(SQL2) <> 0 Then
        
            Set frmMens3 = New frmMensajes
            
            frmMens3.Opcion = 53
            frmMens3.Parametros = SQL2
            frmMens3.Show vbModal
            
            Set frmMens = Nothing
            
            If CadenaDesdeOtroForm <> "OK" Then
                cmdCancelar_Click (0)
            End If
        
        End If
        
        SQL = SQL & " and (cobros.impcobro is null or cobros.impcobro = 0)"
        
    End If
     
     
     
     
     
    Screen.MousePointer = vbHourglass
    Set RS = New ADODB.Recordset
    
    'Marzo 2015
    'Ver si entre los desde hastas hay importes negativos... ABONOS
    
    If SubTipo = vbTipoPagoRemesa Then
    
        'Vemos las cuentas que vamos a girar . Sacaremos codmacta
        Cad = SQL
        Cad = "cobros.codmacta=cuentas.codmacta AND (siturem is null) AND " & Cad
        Cad = Cad & " AND cobros.codforpa = formapago.codforpa ORDER BY codmacta,numfactu "
        Cad = "Select distinct cobros.codmacta FROM cobros,cuentas,formapago WHERE " & Cad
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set colCtas = New Collection
        While Not RS.EOF
            colCtas.Add CStr(RS!codmacta)
            RS.MoveNext
        Wend
        RS.Close
        
        'Ahora veremos los negativos, de las cuentas que vamos a girar
        'Sol el select de los negativos , sin numserie ni na de na
        Cad = "(impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) < 0"
        Cad = "cobros.codmacta=cuentas.codmacta AND (siturem is null) AND " & Cad
        Cad = Cad & " AND cobros.codforpa = formapago.codforpa  "
        Cad = Cad & " AND cobros.situacion = 0 " '++
        Cad = "Select cobros.codmacta,nommacta,numserie,numfactu,impvenci FROM cobros,cuentas,formapago WHERE " & Cad
        
        
        If colCtas.Count > 0 Then
            Cad = Cad & " AND cobros.codmacta IN ("
            For I = 1 To colCtas.Count
                If I > 1 Then Cad = Cad & ","
                Cad = Cad & "'" & colCtas.Item(I) & "'"
            Next
            Cad = Cad & ") ORDER BY codmacta,numfactu"
        
            'Seguimos
        
            Set colCtas = Nothing
            RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Cad = ""
            I = 0
            Set colCtas = New Collection
            While Not RS.EOF
                If I < 15 Then
                    Cad = Cad & vbCrLf & RS!codmacta & " " & RS!Nommacta & "  " & RS!NUmSerie & Format(RS!NumFactu, "000000") & "   -> " & Format(RS!ImpVenci, FormatoImporte)
                End If
                I = I + 1
                colCtas.Add CStr(RS!codmacta)
                RS.MoveNext
            Wend
            RS.Close
            
            If Cad <> "" Then
                If Me.chkComensaAbonos.Value = 0 Then
                
                    If I >= 15 Then Cad = Cad & vbCrLf & "....  y " & I & " vencimientos más"
                    Cad = "Clientes con abonos. " & vbCrLf & Cad & " ¿Continuar?"
                    If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then
                        Set RS = Nothing
                        Set colCtas = Nothing
                        Exit Sub
                    End If
                            
                Else
                    '-------------------------------------------------------------------------
                    CadenaDesdeOtroForm = ""
                    For I = 1 To colCtas.Count
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "'" & colCtas.Item(I) & "',"
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
    I = 0
    If SubTipo = vbTipoPagoRemesa Then
        Cad = " FROM cobros,formapago,cuentas WHERE cobros.codforpa = formapago.codforpa AND (siturem is null) AND situacion = 0 and "
        Cad = Cad & " cobros.codmacta=cuentas.codmacta AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(txtFecha(4).Text), FormatoFecha) & "') AND "
        Cad = "Select cobros.codmacta,nommacta,fecbloq" & Cad & SQL & " GROUP BY 1 ORDER BY 1"
        
    Else
    End If
    
    
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Cad = ""
        I = 1
        While Not RS.EOF
            Cad = Cad & RS!codmacta & " - " & RS!Nommacta & " : " & RS!FecBloq & vbCrLf
            RS.MoveNext
        Wend
    End If

    RS.Close
    
    If I > 0 Then
        Cad = "Las siguientes cuentas estan bloqueadas." & vbCrLf & String(60, "-") & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    
    If SubTipo = vbTipoPagoRemesa Then
        'Efectos bancario
    
        Cad = " FROM cobros,formapago,cuentas WHERE cobros.codforpa = formapago.codforpa AND (siturem is null) AND "
        Cad = Cad & " cobros.codmacta=cuentas.codmacta AND situacion = 0 and "
    Else
    End If
    'Hacemos un conteo
    RS.Open "SELECT Count(*) " & Cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        I = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    Cad = Cad & SQL
    
    
    
    If I > 0 Then
        If SubTipo <> vbTipoPagoRemesa Then
        End If
        I = 1  'Para que siga por abajo
    End If
    
    

    'La suma
    If I > 0 Then
        SQL = "select sum(impvenci),sum(impcobro),sum(gastos) " & Cad
        Impor = 0
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then Impor = DBLet(RS.Fields(0), "N") - DBLet(RS.Fields(1), "N") + DBLet(RS.Fields(2), "N")
        RS.Close
        If Impor = 0 Then I = 0
    End If
        

    Set RS = Nothing
    
    If I = 0 Then
        MsgBox "Ningun dato a remesar con esos valores", vbExclamation
    Else
         
        'Preparamos algunas cosillas
        'Aqui guardaremos cuanto llevamos a cada banco
        SQL = "Delete from tmpCierre1 where codusu =" & vUsu.Codigo
        Conn.Execute SQL
        
        CadenaDesdeOtroForm = ""
        
        'Si son talones o pagares NO hay reajuste en bancos
        'Con lo cual cargare la tabla con el banco
        
        If SubTipo <> vbTipoPagoRemesa Then
            ' Metermos cta banco, nºremesa . El resto no necesito
            SQL = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES ("
            SQL = SQL & vUsu.Codigo & ",'" & txtCuentas(2).Text & "','"
            SQL = SQL & txtRemesa.Text & "',0)"
            Conn.Execute SQL
        Else
            If Not chkAgruparRemesaPorEntidad.Visible Then Me.chkAgruparRemesaPorEntidad.Value = 0
            SQL = Cad 'Le paso el SELECT
            If Me.chkAgruparRemesaPorEntidad.Value = 1 Then
                'lo que yo hacia
                DividiVencimentosPorEntidadBancaria
                
                CadenaDesdeOtroForm = ""
                
                Set frmMens2 = New frmMensajes
                frmMens2.Opcion = 52
                frmMens2.Parametros = txtCuentas(2).Text
                frmMens2.Show vbModal
                Set frmMens2 = Nothing

                If CadenaDesdeOtroForm <> "" Then
                    'Cargamos los vtos
                    PonerVtosRemesa SQL, True
                
                
                    If GenerarRemesa(0) Then
                        MsgBox "Remesa generada correctamente.", vbExclamation
                        cmdCancelar_Click (0)
                        CargaList
                    End If
                Else
                    cmdCancelar_Click (0)
                End If
                
'                PonerVtosRemesa SQL, True
            Else
                'SQL = SQL & VencimientosPorEntidadBancaria   'QUITAR
                
                CargarDatos SQL, True
                
                Dim CadAux As String
                
                CadAux = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES (" & vUsu.Codigo
                CadAux = CadAux & ",'" & txtCuentas(2).Text & "','" & txtNCuentas(2).Text & "'," & DBSet(Text1(4).Text, "N") & ")"
                If Not Ejecuta(CadAux) Then Exit Sub
                
                CadenaDesdeOtroForm = "'" & Trim(txtCuentas(2).Text) & "'"
                
            End If
                                
        End If
        
        
        
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgSerie"
        imgSerie_Click Indice
    Case "imgFecha"
        imgFec_Click Indice
    Case "imgCuentas"
        imgCuentas_Click Indice
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
        Case 0, 1, 2, 3 'cuentas
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
Dim NumeroDocumento As String
Dim CuentasCC As String

    Set miRsAux = New ADODB.Recordset
    
    Conn.Execute "DELETE FROM tmp347 WHERE codusu = " & vUsu.Codigo
    '                                                               POR SI TUVIERAN MISMO BANCO, <> cta contable
    
    
    NumeroDocumento = "select mid(iban,5, 4)  from bancos where not sufijoem is null "
    NumeroDocumento = NumeroDocumento & " and mid(iban,5, 4) > 0  and codmacta<>'" & Me.txtCuentas(2).Text & "' group by 1"
    miRsAux.Open NumeroDocumento, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroDocumento = ""
    While Not miRsAux.EOF
        NumeroDocumento = NumeroDocumento & ", " & miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NumeroDocumento = "" Then
        NumeroDocumento = "-1"
    Else
        NumeroDocumento = Mid(NumeroDocumento, 2) 'quitamos la primera coma
    End If
    
    NumeroDocumento = " (mid(cobros.iban,5, 4)) in (" & NumeroDocumento & ")"
    
    'Agrupamos los vencimientos por entidad,oficina menos los del banco por defecto
    CuentasCC = "select mid(cobros.iban,5, 4) ,sum(impvenci + coalesce(gastos,0)) " & SQL     'FALTA### VER impcobro
    CuentasCC = CuentasCC & " AND " & NumeroDocumento & " GROUP BY 1"
    miRsAux.Open CuentasCC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CuentasCC = "insert into `tmpcierre1` (`codusu`,`cta`,`nomcta`,`acumPerD`) VALUES (" & vUsu.Codigo & ","
        CuentasCC = CuentasCC & miRsAux.Fields(0) & ",0," & TransformaComasPuntos(CStr(miRsAux.Fields(1))) & ")"
        Conn.Execute CuentasCC
        
         miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Los del banco por defecto, y lo que no tenemos banco, es decir, el resto
    '------------------------------------------------------------------------------
    CuentasCC = SQL & " AND ( NOT " & NumeroDocumento & " OR cobros.iban is null) GROUP BY 1"
    'Vere la entidad y la oficina del PPAL
    NumeroDocumento = DevuelveDesdeBD("mid(cobros.iban,5, 4)", "bancos", "codmacta", txtCuentas(2).Text, "T")
    NumeroDocumento = "Select " & NumeroDocumento & ",sum(impvenci + coalesce(gastos,0)) " & CuentasCC      'FALTA### VER impcobro
    miRsAux.Open NumeroDocumento, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CuentasCC = "insert into `tmpcierre1` (`codusu`,`cta`,`nomcta`,`acumPerD`) VALUES (" & vUsu.Codigo & ","
        CuentasCC = CuentasCC & miRsAux.Fields(0) & "," & DBSet(txtNCuentas(2).Text, "T") & "," & DBSet(miRsAux.Fields(1), "N") & ")"
        Conn.Execute CuentasCC
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    espera 1
    
    
    'Pongo codmacta y nombanco como corresponde
    CuentasCC = "Select * from tmpcierre1 where codusu =" & vUsu.Codigo
    miRsAux.Open CuentasCC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        NumeroDocumento = "nommacta"
        CuentasCC = "bancos.codmacta=cuentas.codmacta AND sufijoem<>''  AND mid(bancos.iban,5,4) = " & miRsAux!Cta & " AND 1 "    'ctabancaria.oficina "
        CuentasCC = DevuelveDesdeBD("bancos.codmacta", "bancos,cuentas", CuentasCC, "1", "N", NumeroDocumento)  'miRsAux!nomcta
        If CuentasCC <> "" Then
            CuentasCC = "UPDATE tmpcierre1 SET cta = '" & CuentasCC & "',nomcta ='" & DevNombreSQL(NumeroDocumento)
            CuentasCC = CuentasCC & "' WHERE Cta = '" & miRsAux!Cta & "' AND nomcta =" & DBSet(miRsAux!nomcta, "T")
            Conn.Execute CuentasCC
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Por si quiere borrar alguno de los repartios que hace
    'Por si casao luego BORRAN la remesa a generar para ese banco, es decir , no uqieren llevarlo ahora
    CuentasCC = "insert into tmp347(codusu,cta) select codusu,cta from tmpcierre1 WHERE codusu =" & vUsu.Codigo
    Conn.Execute CuentasCC
    
eDividir:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        
        
    End If
    NumeroDocumento = ""
    CuentasCC = ""
    Set miRsAux = Nothing
    Set RS = Nothing
End Sub

Private Function VencimientosPorEntidadBancaria() As String
Dim SQL As String

    VencimientosPorEntidadBancaria = ""

    SQL = " and length(cobros.iban) <> 0 and mid(cobros.iban,5,4) = (select mid(iban,5,4) from bancos where codmacta = " & DBSet(txtCuentas(2).Text, "T") & ")"
    
    VencimientosPorEntidadBancaria = SQL

End Function


Private Function EfectuarRemesa(Opcion As Integer) As Boolean
Dim C As String
Dim NumeroRemesa As Long
Dim RS As ADODB.Recordset
Dim J As Integer
Dim I As Integer
Dim ImporteQueda As Currency
Dim CodRem As Currency
Dim AnyoRem As Currency
Dim TipoRemesa As Integer
Dim R As ADODB.Recordset

    On Error GoTo EEfectuarRemesa
    EfectuarRemesa = False
    '---------------------------------------------------
    'Creamos la remesa
    SQL = "Select nomcta as numeroremesa,cta from tmpCierre1 where codusu = " & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Datos remesa NO encontrados.", vbExclamation
        miRsAux.Close
        Exit Function
    End If
    If Opcion = 0 Then
        CodRem = miRsAux!NumeroRemesa
        AnyoRem = Year(CDate(txtFecha(4).Text))
    End If
    If cmbRemesa.ListIndex = 1 Then
        TipoRemesa = 3
    Else
        TipoRemesa = 2
    End If

    
    'Si estamos modificando la remesa tenemos que quitar la marca de remeados
    If Opcion = 1 Then
        SQL = "UPDATE  cobros SET siturem= NULL,codrem= NULL, anyorem =NULL,tiporem = NULL"
        SQL = SQL & " WHERE codrem = " & CodRem & " and anyorem =" & AnyoRem
        Conn.Execute SQL
    End If

    Set R = New ADODB.Recordset
    
    'Updateamos los vencimientos.  Desde el listview2 vemos que documentos esta llevando al banco
    For NumRegElim = 1 To lwCobros.ListItems.Count
        
            If lwCobros.ListItems(NumRegElim).Checked Then
                SQL = "Select cobros.numserie, cobros.numfactu, cobros.fecfactu, cobros.fecvenci, talones_facturas.numorden,impvenci, gastos, impcobro, talones.numeroref reftalonpag, codrem, anyorem  "
                SQL = SQL & " FROM (talones inner join talones_facturas on talones.codigo = talones_facturas.codigo) left join cobros on cobros.numserie=talones_facturas.numserie AND cobros.numfactu=talones_facturas.numfactu and"
                SQL = SQL & " cobros.fecfactu=talones_facturas.fecfactu AND cobros.numorden=talones_facturas.numorden"
                SQL = SQL & " WHERE talones.codigo= " & lwCobros.ListItems(NumRegElim).Text
    
                R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not R.EOF
                    SQL = R!NUmSerie 'para que de el error si no existe
                    
                    
                    'La situacion entra directamente a cancelacion cliente
                    SQL = "UPDATE  cobros SET siturem= 'F',codrem= " & CodRem & ", anyorem =" & AnyoRem & ","
                    SQL = SQL & " tiporem = " & TipoRemesa

                    'ponemos la cuenta de banco donde va remesado
                    SQL = SQL & ", ctabanc1 ='" & miRsAux!Cta & "' "
                    'Por si acaso a puesto talon referencia
                    SQL = SQL & " WHERE numserie = '" & R!NUmSerie & "' and numfactu = "
                    SQL = SQL & R!NumFactu & " and fecfactu ='" & Format(R!FecFactu, FormatoFecha)
                    SQL = SQL & "' AND numorden =" & R!numorden
                                
                    Conn.Execute SQL
                    R.MoveNext
                Wend
                R.Close
            End If

    Next NumRegElim


    Importe = Text1(4).Text

    'Cremos la cabecera de las remesas
    If Opcion = 0 Then
        SQL = "insert into `remesas` (`codigo`,`anyo`,`fecremesa`,`situacion`,`codmacta`,`tipo`,`importe`,`descripcion`,`tiporem`) values ("
        SQL = SQL & miRsAux!NumeroRemesa & "," & Year(CDate(txtFecha(4).Text)) & ",'" & Format(txtFecha(4).Text, FormatoFecha) & "','F','"
        SQL = SQL & miRsAux.Fields!Cta & "',NULL," & DBSet(Importe, "N") & ",'" & DevNombreSQL(txtRemesa.Text) & "'," & TipoRemesa & ")"
    Else
        'Updatemaos
        SQL = "UPDATE remesas set importe=" & DBSet(Importe, "N")
        SQL = SQL & ", descripcion = '" & DevNombreSQL(txtRemesa.Text) & "'"
        SQL = SQL & " WHERE codigo = " & CodRem & " AND anyo = " & AnyoRem
    End If
    Conn.Execute SQL

    'Marco en scarecepdoc el llevada a banco
     For NumRegElim = 1 To lwCobros.ListItems.Count
        If lwCobros.ListItems(NumRegElim).Checked Then
        
            SQL = "UPDATE talones SET  LlevadoBanco = 1 WHERE codigo = " & lwCobros.ListItems(NumRegElim).Text
            Conn.Execute SQL
        End If
    Next NumRegElim
    miRsAux.Close
    
    EfectuarRemesa = True
    Set R = Nothing
    Exit Function
EEfectuarRemesa:
    MuestraError Err.Number, Err.Description
    Set R = Nothing
End Function

Private Sub CargaFiltros()
Dim AUX As String
    
    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Pagarés "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "Talones "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2

End Sub

Private Sub NuevaRemTalPag()
Dim CtaPuente As Boolean
Dim ForPa As String
Dim Cad As String
Dim Impor As Currency

'Algunas conideraciones

    
    If Me.cmbRemesa.ListIndex = 0 Then
        CtaPuente = vParamT.PagaresCtaPuente
    Else
        CtaPuente = vParamT.TalonesCtaPuente
    End If



    'A partir de la fecha generemos leemos k remesa corresponde
    SQL = "select max(codigo) from remesas where anyo=" & Year(CDate(txtFecha(4).Text))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then
        NumRegElim = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    NumRegElim = NumRegElim + 1
    txtRemesa.Tag = NumRegElim

    
    If Me.cmbRemesa.ListIndex = 0 Then
        SQL = " talon = 0"
    Else
        SQL = " talon = 1"
    End If

    'Si no lleva cuenta puente, no hace falta que este contabilizada
    'Es decir. Solo mirare contabilizados si llevo ctapuente
    If CtaPuente Then SQL = SQL & " AND contabilizada = 1 "
    SQL = SQL & " AND LlevadoBanco = 0 "
    
    'de la recepcion de factura, fecha de vencimiento
    If txtFecha(2).Text <> "" Then SQL = SQL & " AND fechavto >= '" & Format(txtFecha(2).Text, FormatoFecha) & "'"
    If txtFecha(3).Text <> "" Then SQL = SQL & " AND fechavto <= '" & Format(txtFecha(3).Text, FormatoFecha) & "'"

    'Fecha recepcion
    If txtFecha(0).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    If txtFecha(1).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    
    
    Screen.MousePointer = vbHourglass
    Set RS = New ADODB.Recordset
    
    'Que la cuenta NO este bloqueada
    I = 0
    Cad = "select cuentas.codmacta,nommacta,FecBloq from "
    Cad = Cad & "talones,cuentas where talones.codmacta=cuentas.codmacta"
    Cad = Cad & " AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(txtFecha(4).Text), FormatoFecha) & "') "
    Cad = Cad & " AND " & SQL & " GROUP by 1"

    
    
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Cad = ""
        I = 1
        While Not RS.EOF
            Cad = Cad & RS!codmacta & " - " & RS!Nommacta & " : " & RS!FecBloq & vbCrLf
            RS.MoveNext
        Wend
    End If

    RS.Close
    
    If I > 0 Then
        Cad = "Las siguientes cuentas estan bloquedas." & vbCrLf & String(60, "-") & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    

    Cad = " FROM talones,cuentas where talones.codmacta=cuentas.codmacta AND"

    'Hacemos un conteo
    RS.Open "SELECT * " & Cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    Cad = ""
    While Not RS.EOF
        I = I + 1
        Cad = Cad & " OR ( codigo = " & RS!Codigo & ") "
        RS.MoveNext
    Wend
    RS.Close
    If I = 0 Then
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Sub
    End If
    Cad = "(" & Mid(Cad, 4) & ")"
    SQL = " from cobros where (numserie,numfactu,fecfactu,numorden) in (select numserie ,numfactu,fecfactu,numorden from talones_facturas where " & Cad & ")"
    SQL = "select sum(impvenci),sum(impcobro),sum(gastos) " & SQL
    
    
    

    'La suma
    If I > 0 Then
        
        Impor = 0
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'If Not Rs.EOF Then Impor = DBLet(Rs.Fields(0), "N") - DBLet(Rs.Fields(1), "N") + DBLet(Rs.Fields(2), "N")
        
        'Solo el impcobro
        If Not RS.EOF Then Impor = DBLet(RS.Fields(1), "N")
        RS.Close
        If Impor = 0 Then I = 0
    End If
        

    Set RS = Nothing
    
    If I = 0 Then
        MsgBox "Ningun dato a remesar con esos valores(II)", vbExclamation
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
            SQL = SQL & txtRemesa.Tag & "',0)"
            Conn.Execute SQL
        End If
        
        
        'Lo qu vamos a hacer es , primero bloquear la opcioin de remesar
        If BloqueoManual(True, "Remesas", "Remesas") Then
            
'            Me.Visible = False
           
            
            CargarDatos Cad, True
            

'            'Remesas de talones y pagares
'            frmRemeTalPag.vRemesa = "" 'NUEVA
'            frmRemeTalPag.SQL = Cad
'            frmRemeTalPag.Talon = cmbRemesa.ListIndex = 1 '0 pagare   1 talon
'            frmRemeTalPag.Text1(0).Text = Me.txtCta(3).Text & " - " & txtDescCta(3).Text
'            frmRemeTalPag.Text1(1).Text = Text1(8).Text
'            frmRemeTalPag.Show vbModal

            'Desbloqueamos
            BloqueoManual False, "Remesas", ""
            
        Else
            MsgBox "Otro usuario esta generando remesas", vbExclamation
        End If

    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Function ComprobarEfectosCobradosParcialmente() As Boolean
Dim AUX As String
Dim MasDeUnDocumento As Byte
    On Error GoTo EComprobarEfectosCobradosParcialmente
    ComprobarEfectosCobradosParcialmente = False
    
    
        AUX = ""
        MasDeUnDocumento = 0
        For NumRegElim = 1 To ListView2.ListItems.Count
            If lwCobros.ListItems(NumRegElim).Checked Then
                'Este documento. Vemos el importe del documento
                AUX = AUX & "," & lwCobros.ListItems(NumRegElim).Text
                If MasDeUnDocumento = 0 Then
                    MasDeUnDocumento = 1
                Else
                    MasDeUnDocumento = 2
                End If
            End If
        Next
        
        AUX = Mid(AUX, 2) 'quito la primera coma
        If MasDeUnDocumento > 1 Then
            
            '1. Si existe algun vto cobrado parcialmente y recepcionado en dos de los documentos que vamos a recepcionar
            SQL = "Select cobros.numserie,cobros.numfactu,cobros.fecfactu,cobros.numorden,count(*)"
            SQL = SQL & " FROM talones_facturas left join cobros on cobros.numserie=talones_facturas.numserie AND cobros.numfactu=talones_facturas.numfactu and"
            SQL = SQL & " cobros.fecfactu = talones_facturas.fecfactu And cobros.numorden = talones_facturas.numorden"
            SQL = SQL & " WHERE codigo in (" & AUX & ") group by 1,2,3,4 having count(*) >1"
        
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            While Not miRsAux.EOF
                SQL = SQL & miRsAux!NUmSerie & miRsAux!NumFactu & " / " & miRsAux!numorden & vbCrLf
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        
            If SQL <> "" Then
                SQL = "Los siguientes vencimientos estan mas de una vez: " & vbCrLf & SQL & vbCrLf
                SQL = SQL & "No deberia seguir con el proceso. ¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
                'Exit Sub  'FALTA### ver si hay que salir
            End If
        End If
        
        'Veremos si los vtos estan ya remesados
        SQL = "Select cobros.numserie,cobros.numfactu,cobros.fecfactu,cobros.numorden"
        SQL = SQL & " FROM talones_facturas left join cobros on cobros.numserie=talones_facturas.numserie AND cobros.numfactu=talones_facturas.numfactu and"
        SQL = SQL & " cobros.fecfactu = talones_facturas.fecfactu And cobros.numorden = talones_facturas.numorden and codrem>0"
        SQL = SQL & " WHERE codigo in (" & AUX & ") group by 1,2,3,4"
        
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            If Not IsNull(miRsAux!NUmSerie) Then SQL = SQL & miRsAux!NUmSerie & miRsAux!NumFactu & " / " & miRsAux!numorden & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
        If SQL <> "" Then
            SQL = "Los siguientes vencimientos estan remesados y no ha sido eliminado el riesgo: " & vbCrLf & SQL
            MsgBox SQL, vbExclamation
            Exit Function
        End If
        
    ComprobarEfectosCobradosParcialmente = True
    
    Exit Function
EComprobarEfectosCobradosParcialmente:
    MuestraError Err.Number, Err.Description
End Function

