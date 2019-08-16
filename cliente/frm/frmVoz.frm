VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVoz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TELMEX Telefonía"
   ClientHeight    =   10530
   ClientLeft      =   630
   ClientTop       =   525
   ClientWidth     =   16410
   Icon            =   "frmVoz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   16410
   Tag             =   "0"
   Begin VB.Frame fraElementosFacturar 
      BackColor       =   &H00C09258&
      Caption         =   "  Información actual del producto  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2310
      Width           =   16365
   End
   Begin Threed.SSPanel pnlTLinea 
      Height          =   1815
      Left            =   60
      TabIndex        =   62
      Top             =   2550
      Width           =   1245
      _Version        =   65536
      _ExtentX        =   2196
      _ExtentY        =   3201
      _StockProps     =   15
      Caption         =   "Corporativa"
      BackColor       =   6999100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      MousePointer    =   99
      MouseIcon       =   "frmVoz.frx":0CCA
      Begin VB.Label lblTLinea 
         Alignment       =   1  'Right Justify
         BackColor       =   &H006ACC3C&
         Caption         =   "Tipos Linea "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         MouseIcon       =   "frmVoz.frx":15A4
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   90
         Width           =   1065
      End
      Begin VB.Image imgTlinea 
         Height          =   1455
         Left            =   90
         MouseIcon       =   "frmVoz.frx":1E6E
         Picture         =   "frmVoz.frx":2738
         Top             =   270
         Width           =   1065
      End
      Begin VB.Image imgTLineaDark 
         Height          =   1455
         Left            =   90
         MouseIcon       =   "frmVoz.frx":30FB
         MousePointer    =   99  'Custom
         Picture         =   "frmVoz.frx":39C5
         Top             =   270
         Width           =   1065
      End
   End
   Begin Threed.SSPanel pnlPublica 
      Height          =   1485
      Left            =   0
      TabIndex        =   65
      Top             =   5310
      Width           =   1545
      _Version        =   65536
      _ExtentX        =   2725
      _ExtentY        =   2619
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      MousePointer    =   99
      MouseIcon       =   "frmVoz.frx":8BDF
      Begin VB.Label lblPublica 
         Alignment       =   1  'Right Justify
         Caption         =   "TPBCL          "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   66
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Image imgPublicaDark 
         Height          =   1110
         Left            =   90
         MouseIcon       =   "frmVoz.frx":94B9
         MousePointer    =   99  'Custom
         Picture         =   "frmVoz.frx":9D83
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1365
      End
      Begin VB.Image imgPublica 
         Height          =   1110
         Left            =   90
         MouseIcon       =   "frmVoz.frx":11421
         Picture         =   "frmVoz.frx":11CEB
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1365
      End
   End
   Begin Threed.SSPanel pnlCorporativa 
      Height          =   1485
      Left            =   180
      TabIndex        =   64
      Top             =   4020
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   2619
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      MousePointer    =   99
      MouseIcon       =   "frmVoz.frx":19389
      Begin VB.Label lblCorporativa 
         Alignment       =   1  'Right Justify
         Caption         =   "Corporativa "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   67
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Image imgCorporativaDark 
         Height          =   1110
         Left            =   90
         MouseIcon       =   "frmVoz.frx":19C63
         MousePointer    =   99  'Custom
         Picture         =   "frmVoz.frx":1A52D
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1305
      End
      Begin VB.Image imgCorporativa 
         Height          =   1110
         Left            =   90
         MouseIcon       =   "frmVoz.frx":1FADB
         Picture         =   "frmVoz.frx":203A5
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1305
      End
   End
   Begin VB.Frame FraDatosGenerales 
      BackColor       =   &H00C09258&
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   16365
   End
   Begin VB.Frame fraDatosIncidenteUltimo 
      BackColor       =   &H00C09258&
      Caption         =   " Asuntos que han modificado la información  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -60
      TabIndex        =   16
      Top             =   7110
      Width           =   16425
   End
   Begin VB.Frame fraNombreCliente 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   16365
      Begin VB.TextBox txtIDCliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1230
         Width           =   2265
      End
      Begin VB.TextBox txtNombreCliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4110
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1230
         Width           =   5655
      End
      Begin VB.TextBox txtIdDatosVoz 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   900
         Width           =   2265
      End
      Begin VB.Frame FraEnlace 
         Enabled         =   0   'False
         Height          =   915
         Left            =   0
         TabIndex        =   1
         Top             =   -30
         Width           =   16335
         Begin VB.Frame Frame2 
            Height          =   705
            Left            =   12930
            TabIndex        =   28
            Top             =   120
            Width           =   615
            Begin VB.Image Image1 
               Height          =   480
               Index           =   1
               Left            =   60
               Picture         =   "frmVoz.frx":20E4F
               Top             =   150
               Width           =   480
            End
         End
         Begin VB.TextBox txtIdProducto 
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
            Left            =   1830
            TabIndex        =   4
            Top             =   510
            Width           =   2265
         End
         Begin VB.TextBox txtCodigoEnlace 
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
            Left            =   1830
            TabIndex        =   3
            Top             =   180
            Width           =   2265
         End
         Begin VB.TextBox txtNombreProducto 
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
            Left            =   4080
            TabIndex        =   2
            Top             =   510
            Width           =   8475
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   135
            Left            =   10890
            TabIndex        =   29
            Top             =   300
            Width           =   5385
            _Version        =   65536
            _ExtentX        =   9499
            _ExtentY        =   238
            _StockProps     =   15
            BackColor       =   14063626
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblIdProducto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C09258&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Id del Producto:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   150
            TabIndex        =   8
            Top             =   510
            Width           =   1665
         End
         Begin VB.Label lblCodigoEnlace 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C09258&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Código de Enlace:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   150
            TabIndex        =   7
            Top             =   180
            Width           =   1665
         End
         Begin VB.Label lblColor1 
            BackColor       =   &H00C09258&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   5910
            TabIndex        =   6
            Top             =   180
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.Label lblColor2 
            BackColor       =   &H00FDEFDF&
            BorderStyle     =   1  'Fixed Single
            Height          =   165
            Left            =   6180
            TabIndex        =   5
            Top             =   180
            Visible         =   0   'False
            Width           =   165
         End
      End
      Begin VB.Label lblIDCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID del Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   13
         Top             =   1290
         Width           =   930
      End
      Begin VB.Label lblIdFacturacion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id de Datos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   510
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame fraIncidenteActual 
      Height          =   2970
      Left            =   -30
      TabIndex        =   17
      Top             =   7095
      Width           =   16395
      Begin VB.CommandButton cmdEditarVoz 
         Caption         =   "Agregar Nuevo Asunto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6600
         TabIndex        =   63
         ToolTipText     =   "Agregar un nuevo asunto que modifica la información"
         Top             =   2610
         Width           =   3615
      End
      Begin VB.CommandButton cmdVerAnteriores 
         Caption         =   "Editar Asunto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10245
         TabIndex        =   18
         ToolTipText     =   "Actualización por el incidente seleccionado"
         Top             =   2610
         Width           =   3570
      End
      Begin MSFlexGridLib.MSFlexGrid grdAsuntosModificaron 
         Height          =   2355
         Left            =   45
         TabIndex        =   19
         Top             =   270
         Width           =   16245
         _ExtentX        =   28654
         _ExtentY        =   4154
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   375
         Left            =   30
         TabIndex        =   20
         Top             =   2535
         Width           =   3615
         Begin VB.Label Label1 
            Caption         =   "Ultimo asunto que modifico la Información"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   420
            TabIndex        =   23
            Top             =   120
            Width           =   3165
         End
         Begin VB.Label lblDespintar 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   165
            Left            =   3390
            TabIndex        =   22
            Top             =   210
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.Label lblPintar 
            BackColor       =   &H00FDEFDF&
            BorderStyle     =   1  'Fixed Single
            Height          =   165
            Left            =   150
            TabIndex        =   21
            Top             =   150
            Width           =   165
         End
      End
   End
   Begin VB.Frame fraBotones 
      Height          =   555
      Left            =   0
      TabIndex        =   24
      Top             =   9960
      Width           =   16365
      Begin VB.CommandButton cmdLogCanales 
         Caption         =   "Log Canales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4970
         TabIndex        =   98
         Top             =   180
         Width           =   1710
      End
      Begin VB.CommandButton cmdAprobacionNumeros 
         Caption         =   "Aprobación Numeros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   180
         Width           =   1710
      End
      Begin VB.CommandButton cmdTickets 
         Caption         =   "Tickets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   180
         Width           =   1560
      End
      Begin VB.CommandButton cmdAdministracion 
         Caption         =   "Administración"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   180
         Width           =   1560
      End
      Begin VB.CommandButton cmdVerInterfase 
         Caption         =   "&Ver Interfase pendiente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   13650
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   210
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin Threed.SSPanel pnlCliente 
      Height          =   795
      Left            =   30
      TabIndex        =   34
      Top             =   1530
      Width           =   16335
      _Version        =   65536
      _ExtentX        =   28813
      _ExtentY        =   1402
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel2 
         Height          =   285
         Left            =   0
         TabIndex        =   60
         Top             =   225
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Cliente de Telefonia Nacional"
         ForeColor       =   16777215
         BackColor       =   12620376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdELiminarClienteLocal 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7860
         TabIndex        =   36
         Top             =   1770
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdBuscarClienteLocal 
         Caption         =   "&Buscar..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7860
         TabIndex        =   35
         Top             =   1500
         Visible         =   0   'False
         Width           =   1185
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   285
         Left            =   9840
         TabIndex        =   85
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Telefonia Local"
         ForeColor       =   16777215
         BackColor       =   12620376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblEstrato 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   96
         Top             =   510
         Width           =   3135
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Estrato :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   17
         Left            =   7110
         TabIndex        =   95
         Top             =   570
         Width           =   705
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Enlace:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   11160
         TabIndex        =   89
         Top             =   540
         Width           =   765
      End
      Begin VB.Label LblEnlace 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12000
         TabIndex        =   88
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Id Venta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   11160
         TabIndex        =   87
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LblVenta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12000
         TabIndex        =   86
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lblTItulo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "TELEFONIA NACIONAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   59
         Top             =   30
         Width           =   9000
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "ID Cliente :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   58
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblIDCliente1 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   57
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   56
         Top             =   570
         Width           =   765
      End
      Begin VB.Label lblCliente 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   55
         Top             =   510
         Width           =   1635
      End
      Begin VB.Label lblClienteLocal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   54
         Top             =   1800
         Visible         =   0   'False
         Width           =   4050
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B98D1C&
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   30
         TabIndex        =   53
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblIDClienteLocal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   52
         Top             =   1500
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B98D1C&
         Caption         =   "ID Cliente :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   30
         TabIndex        =   51
         Top             =   1530
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblTItulo 
         Caption         =   "Indique aquí el cliente que tiene o tendrá registrado el producto de telefonía local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   50
         Top             =   1290
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Label lblTItulo 
         Alignment       =   2  'Center
         BackColor       =   &H00B98D1C&
         Caption         =   "TELEFONIA LOCAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   30
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   9000
      End
      Begin VB.Label lblCiudad 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   48
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Ciudad :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   47
         Top             =   255
         Width           =   690
      End
      Begin VB.Label lblDireccion 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   46
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   7080
         TabIndex        =   45
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblsede 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5610
         TabIndex        =   44
         Top             =   510
         Width           =   1425
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Sede :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   4920
         TabIndex        =   43
         Top             =   570
         Width           =   585
      End
      Begin VB.Label lblCiudadlocal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3495
         TabIndex        =   42
         Top             =   1530
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B98D1C&
         Caption         =   "Ciudad :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   2565
         TabIndex        =   41
         Top             =   1560
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblDireccionLocal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6015
         TabIndex        =   40
         Top             =   1530
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B98D1C&
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   5085
         TabIndex        =   39
         Top             =   1560
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblSedeLocal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00B98D1C&
         Caption         =   "Sede :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   5100
         TabIndex        =   37
         Top             =   1830
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Frame frmTLinea 
      Height          =   4650
      Left            =   1080
      TabIndex        =   14
      ToolTipText     =   "Agregar un nuevo asunto que modifique la información"
      Top             =   2445
      Width           =   15285
      Begin Threed.SSPanel pnlExplicacion 
         Height          =   375
         Index           =   0
         Left            =   300
         TabIndex        =   74
         Top             =   450
         Width           =   8505
         _Version        =   65536
         _ExtentX        =   15002
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Los tipos de línea corresponden al tipo de conexión física entre la oficina y la red. Estos pueden ser E1, RDSI, Básicas etc."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdDatosVoz 
         Height          =   3405
         Left            =   630
         TabIndex        =   69
         Top             =   720
         Width           =   14565
         _ExtentX        =   25691
         _ExtentY        =   6006
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin Threed.SSPanel pnlVerde 
         Height          =   135
         Left            =   4650
         TabIndex        =   70
         Top             =   4170
         Visible         =   0   'False
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   238
         _StockProps     =   15
         BackColor       =   6999100
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkCancelados 
         Caption         =   "Ver líneas canceladas"
         Height          =   195
         Left            =   10095
         TabIndex        =   30
         Top             =   4305
         Width           =   2025
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   540
         TabIndex        =   31
         Top             =   4050
         Width           =   3615
         Begin VB.Label lblCancelados 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   165
            Left            =   150
            TabIndex        =   33
            Top             =   210
            Width           =   165
         End
         Begin VB.Label Label2 
            Caption         =   "Líneas que se encuentran canceladas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   405
            TabIndex        =   32
            Top             =   180
            Width           =   3165
         End
      End
      Begin Threed.SSPanel pnlRosa 
         Height          =   135
         Left            =   5040
         TabIndex        =   71
         Top             =   4170
         Visible         =   0   'False
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   238
         _StockProps     =   15
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlAmarillo 
         Height          =   135
         Left            =   5430
         TabIndex        =   72
         Top             =   4170
         Visible         =   0   'False
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   238
         _StockProps     =   15
         BackColor       =   10085367
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlGris 
         Height          =   135
         Left            =   5820
         TabIndex        =   73
         Top             =   4170
         Visible         =   0   'False
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   238
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblGrupoCentrex 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10860
         TabIndex        =   93
         Top             =   150
         Width           =   1785
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Grupo Centrex:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   16
         Left            =   9600
         TabIndex        =   92
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label lblCallSource 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10860
         TabIndex        =   91
         Top             =   420
         Width           =   1785
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Call Source:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   15
         Left            =   9600
         TabIndex        =   90
         Top             =   465
         Width           =   1185
      End
   End
   Begin VB.Frame frmPublica 
      Height          =   4635
      Left            =   1080
      TabIndex        =   76
      Top             =   2460
      Width           =   15285
      Begin Threed.SSPanel pnlExplicacion 
         Height          =   375
         Index           =   4
         Left            =   -30
         TabIndex        =   80
         Top             =   180
         Width           =   7065
         _Version        =   65536
         _ExtentX        =   12462
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "La numeración pública corresponde a los números telefónicos asignados al cliente"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdNumeracionPublica 
         Height          =   4005
         Left            =   660
         TabIndex        =   81
         Top             =   480
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   7064
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin Threed.SSPanel pnlExplicacion 
         Height          =   375
         Index           =   5
         Left            =   7290
         TabIndex        =   82
         Top             =   180
         Width           =   5685
         _Version        =   65536
         _ExtentX        =   10028
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Servicios Suplementarios"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdServiciosSuplementarios 
         Height          =   4005
         Left            =   7200
         TabIndex        =   83
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   7064
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdNomenServiciosSuple 
         Height          =   4005
         Left            =   11910
         TabIndex        =   97
         Top             =   480
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   7064
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame frmCorporativa 
      Height          =   4605
      Left            =   1080
      TabIndex        =   75
      Top             =   2460
      Width           =   15285
      Begin Threed.SSPanel pnlExplicacion 
         Height          =   375
         Index           =   1
         Left            =   300
         TabIndex        =   77
         Top             =   180
         Width           =   6645
         _Version        =   65536
         _ExtentX        =   11721
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "La numeración privada corresponde al plan de numeración de extensiones del cliente"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdNumeracionPrivada 
         Height          =   4005
         Left            =   660
         TabIndex        =   78
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7064
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin Threed.SSPanel pnlExplicacion 
         Height          =   375
         Index           =   2
         Left            =   7230
         TabIndex        =   79
         Top             =   180
         Width           =   6765
         _Version        =   65536
         _ExtentX        =   11933
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Plan de numeración actual del cliente"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin MSComctlLib.TreeView trvPlanActual 
         Height          =   3945
         Left            =   7350
         TabIndex        =   84
         Top             =   510
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   6959
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmVoz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Se agrega el boton Log Canales
'       VERSION:       1.0.200
'       REQUERIMIENTO: 3488
'       FECHA:       2008/07/24
'*******************************************************************
'**********************************************************************
' MODIFICADO POR :      CARLOS ALBERTO BARRERA
' DESCRIPCION CAMBIO:   Se pasa como parametro la propiedad del id del cliente
' VERSION: 1.0.100
' FECHA: SEPTIEMBRE 7/2009
'****************************************************************
'**********************************************************************
' MODIFICADO POR :      CARLOS ALBERTO BARRERA
' DESCRIPCION CAMBIO:   Se coloca como tipo Duoble el param
' VERSION: 1.1.000
' FECHA: ABRIL 21/2010
'****************************************************************

Option Explicit

Public proConexion As ADODB.Connection

Public proDatosProductoId As String
Public proCompanyId As String
Public proCompanyName As String


'Propiedad que tiene el id del cliente
Public proiClienteId As Long '1.0.100

Public proOnyx As EDCVoz.claONYX

Public proDatosProducto As claDatosProducto

Public proServiciosSuplementarios As EDCAdminVoz.colServiciosSup


'Colección de Usuarios que pueden ver el Botón de Administración
Dim varColUsuarios As EDCAdminVoz.colUsuario
'Dim varBloqueo As claBloqueo
Dim varPassword As String
'Propiedad para saber desde donde se llamo el load de la forma
'F:     Facturacion
'T:     Ajuste de Tarifas
Public proLlamadoForma As String

Private varPlanNumeracion As colPlanNumeracion

Private varClienteTelefonia As EDCVoz.claClienteTelefonia

Private Sub cmdLogCanales_Click()
    On Error GoTo ErrManager
    
    'Pasar la conexion

    Set frmBuscarCanales.proConexion = Me.proConexion
    frmBuscarCanales.proCompanyId = Me.txtIDCliente
       
       
    'Abrir la ventana de edicion
    frmBuscarCanales.Show (1)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub chkCancelados_Click()
    On Error GoTo ErrManager
    
    Call SubFPintarGridDetalles
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdAdministracion_Click()
Dim varAdministrador As EDCAdminVoz.claONYX
On Error GoTo ErrorManager

    'instancia la Variable
    Set varAdministrador = New EDCAdminVoz.claONYX
    
    'Copia las propiedades a la clase ONYX del Administrador
    varAdministrador.AlternateID = Me.proOnyx.AlternateID
    varAdministrador.ContactID = Me.proOnyx.ContactID
    varAdministrador.ContactName = Me.proOnyx.ContactName
    varAdministrador.DatabaseName = Me.proOnyx.DatabaseName
    varAdministrador.DetailID = Me.proOnyx.DetailID
    varAdministrador.ServerName = Me.proOnyx.ServerName
    varAdministrador.UserLogin = Me.proOnyx.UserLogin
    varAdministrador.UserPassword = Me.proOnyx.UserPassword
    varAdministrador.UserSite = Me.proOnyx.UserSite
    
    varAdministrador.Initiate
    
    'Si es una modificacion
'    If Val(Trim(Me.proAsignacionId)) <> 0 Then
'        'Recuperar la informacion de los incidentes que modificaron la facturacion
'        Me.proDatosProducto.proDatosProductoId = CLng(Me.proAsignacionId)
'
'        Call SubGLLenarColeccionFacturacion
'
'        'Refrescar Informacion del grid de asuntos
'        Call SubFRefrescarGridAsuntos
'
'        'Refrescar Informacion del grid de facturacion
'        Call SubGPintarFacturacionTarifa
'
'        'Recuperar la informacion del enlace y el producto
'        If Not Me.FunFRecuperarDatosGenerales Then
'            MsgBox "Error a Recuperar la información del producto.", vbCritical, App.Title
'        End If
'    End If

    
    Exit Sub
    
ErrorManager:
    SubGMuestraError

End Sub

Private Sub cmdAprobacionNumeros_Click()
On Error GoTo ErrManager
    
    'Pasar la conexion
    Set frmAprobacionNumeros.proConexion = Me.proConexion
    frmAprobacionNumeros.proUserOnyx = Me.proOnyx.UserLogin
 
    'Abrir la ventana de edicion
    frmAprobacionNumeros.Show (1)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdEditarVoz_Click()
    On Error GoTo ErrManager
    

    'Pasar la conexion
    Set frmVozIncident.proConexion = Me.proConexion
    Set frmVozIncident.proDatosProducto = Me.proDatosProducto
    Set frmVozIncident.proOnyx = Me.proOnyx
    Set frmVozIncident.proClienteTelefonia = varClienteTelefonia
    
    frmVozIncident.proInsUpd = "I"
    
    'Abrir la ventana de edicion
    frmVozIncident.Show (1)
    Me.lblGrupoCentrex.Caption = varClienteTelefonia.proGrupoCentrex
    Me.lblCallSource.Caption = varClienteTelefonia.proCallSource
    Me.LblVenta = Me.proDatosProducto.proiVentaid
    Me.lblCodigoEnlace = Me.proDatosProducto.proCodigoEnlace
    
    'Mostrar la información del encabezado
    Me.txtCodigoEnlace.Text = Me.proDatosProducto.proCodigoEnlace
    Me.txtIdProducto.Text = Val(Me.proDatosProducto.proProductId)
    Me.txtNombreProducto.Text = Me.proDatosProducto.proProductName
    
    Me.txtIdDatosVoz.Text = Me.proDatosProductoId
    Me.txtIDCliente.Text = Me.proCompanyId
    Me.txtNombreCliente.Text = Me.proCompanyName
    Me.lblEstrato = Me.proDatosProducto.proDescripcionEstrato
        
    Call SubFInicializarGridDetalle
    
    Call SubFPintarGridDetalles
    
    'Consultar la información de los asuntos
    If Me.proDatosProducto.MetConsultarIncidentes Then
        Call SubFPintarGridIncidentes
    Else
        MsgBox "Error al consultar la información de los incidentes.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdTickets_Click()
    On Error GoTo ErrManager
    
    'Pasar la conexion

    Set frmTicketsEnlace.proConexion = Me.proConexion
    Set frmTicketsEnlace.proOnyx = Me.proOnyx
    Set frmTicketsEnlace.proDatosProducto = Me.proDatosProducto
    frmTicketsEnlace.provchSerialNumber = Me.txtCodigoEnlace.Text
       
    'Abrir la ventana de edicion
    frmTicketsEnlace.Show (1)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdVerAnteriores_Click()
    Dim varProceso As claProceso
    On Error GoTo ErrManager
    
    'Verificación del proceso
    Set varProceso = New claProceso
    Set varProceso.proConexion = Me.proConexion
    
    If Me.proDatosProducto.proDatosProductoIncident Is Nothing Or Me.proDatosProducto.proDatosProductoIncident.Count = 0 Then
        MsgBox "Debe agregar un nuevo incidente modificar la información.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Me.grdAsuntosModificaron.Row = 0 Then
        MsgBox "Debe seleccionar el asunto que desea editar.", vbInformation, App.Title
        Exit Sub
    End If
    
    Me.proDatosProducto.proIncidentId = Me.proDatosProducto.proDatosProductoIncident.Item(Me.grdAsuntosModificaron.Row).proIncidentId
    varProceso.proIncidentId = Me.proDatosProducto.proIncidentId
    varProceso.proUsuario = Me.proOnyx.UserLogin
    varProceso.proNoValidar = varGValidacion
    
    If varProceso.MetValidaPermisos = False Then
        Exit Sub
    End If
    
    If varProceso.proAcceso = AccesoDenegado Then Exit Sub
    
    'Pasar la conexion
    Set frmVozIncident.proConexion = Me.proConexion
    Set frmVozIncident.proDatosProducto = Me.proDatosProducto
    Set frmVozIncident.proOnyx = Me.proOnyx
    Set frmVozIncident.proClienteTelefonia = varClienteTelefonia
    frmVozIncident.proInsUpd = "U"
    
    'Abrir la ventana de edicion
    frmVozIncident.Show (1)
    Me.LblVenta = Me.proDatosProducto.proiVentaid
    Me.lblCodigoEnlace = Me.proDatosProducto.proCodigoEnlace
    Call SubFPintarGridDetalles
    If Me.proDatosProducto.MetConsultar Then
        Me.lblEstrato = Me.proDatosProducto.proDescripcionEstrato
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Activate()
    Dim varSalir As String
    Dim varDatosProductoId As Double '* 1.1.000 Inicio Se cambio el tipo a Double para varDatosProductoId
    On Error GoTo ErrManager
    varSalir = Me.Tag
    If Len(varSalir) = 0 Then
        varSalir = "0"
    End If
    If Len((Trim(Me.proDatosProductoId))) = 0 Then
        varDatosProductoId = 0
    Else
        varDatosProductoId = Val(Trim(Me.proDatosProductoId))
    End If
    
    If varDatosProductoId = 0 And varSalir = "0" Then
        'si es un registro nuevo enviarlo a la ventana de edicion
        Call cmdEditarVoz_Click
    End If
     
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    Dim varColCliente As EDCTraslados.colCliente
    On Error GoTo ErrManager
    
    Set varColCliente = New EDCTraslados.colCliente
    Set varColCliente.proConexion = Me.proConexion
        
    varColCliente.proClienteId = Me.proCompanyId
    If varColCliente.funGConsultaClientexID = True Then
        Me.lblCiudad = varColCliente.Item(1).proCiudad
        Me.lblDireccion = varColCliente.Item(1).proDireccion
        Me.lblsede = varColCliente.Item(1).proSede
        Me.lblIDCliente1 = varColCliente.Item(1).proClienteId
        Me.lblCliente = varColCliente.Item(1).proNombreCliente
        
    Else
            MsgBox "No fue posible encontrar el cliente " & Me.proDatosProducto.proClienteNacionalId
    End If
    
    'Buscar el Grupo Centrex y el Call Source
    Set varClienteTelefonia = New claClienteTelefonia
    Set varClienteTelefonia.proConexion = Me.proConexion
    
    varClienteTelefonia.proCompanyId = Me.proCompanyId
    
    If varClienteTelefonia.MetConsultarxCliente Then
        Me.lblGrupoCentrex.Caption = varClienteTelefonia.proGrupoCentrex
        Me.lblCallSource.Caption = varClienteTelefonia.proCallSource
    Else
        MsgBox "Error al consultar la información del Grupo Centrex y del Call Source.", vbCritical, App.Title
    End If
    
    '-------------
    Call SubFSeguridad
    Call SubFAprobacionNumeros
    Call SubFInicializarGridIncidentes
    Call SubFInicializarGridNumeracionCorporativa
    Call SubFInicializarGridNumeroPublico
    Call SubFInicializarGridServiciosSuplementarios
    Call SubFInicializarGridNomenServiciosSuplementarios
    
    If Me.proDatosProducto Is Nothing Then
        Set Me.proDatosProducto = New claDatosProducto
        Set Me.proDatosProducto.proConexion = Me.proConexion
    End If
    
    Me.proDatosProducto.proDatosProductoId = Me.proDatosProductoId
    
    'Si es una modificacion
    If Val(Trim(Me.proDatosProductoId)) <> 0 Then
        
        If Me.proDatosProducto.MetConsultar Then
            
            'Mostrar la información del encabezado
            Me.txtCodigoEnlace.Text = Me.proDatosProducto.proCodigoEnlace
            Me.txtIdProducto.Text = Me.proDatosProducto.proProductId
            Me.txtNombreProducto.Text = Me.proDatosProducto.proProductName
            Me.txtIdDatosVoz.Text = Me.proDatosProductoId
            Me.txtIDCliente.Text = Me.proCompanyId
            Me.txtNombreCliente.Text = Me.proCompanyName
            Me.LblVenta = Me.proDatosProducto.proiVentaid
            Me.LblEnlace = Me.proDatosProducto.proCodigoEnlace
            Me.lblEstrato = Me.proDatosProducto.proDescripcionEstrato
            
            'Consultar la información de los asuntos
            If Me.proDatosProducto.MetConsultarIncidentes Then
                Call SubFPintarGridIncidentes
            Else
                MsgBox "Error al consultar la información de los incidentes.", vbCritical, App.Title
                Exit Sub
            End If
            
            If Me.proDatosProducto.MetConsultarProductMaster Then
               '* 1.0.100 Inicio Se pasa la propiedad del id del cliente
                If Me.proDatosProducto.MetConsultarParametrosProducto(proiClienteId) Then
                '* 1.0.100 Fin
                    Call SubFInicializarGridDetalle
                Else
                    MsgBox "Error al consultar los datos del producto."
                End If
            Else
                MsgBox "Error al consultar el producto.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Consultar información de los Detalles
            If Me.proDatosProducto.MetConsultarDetalles Then
                Call SubFPintarGridDetalles
            Else
                MsgBox "Error al consultar los detalles del Tab de Datos por Servicios.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Consultar la información de la numeración Privada
            If Me.proDatosProducto.MetConsultarNumeracionCorporativa Then
                Call SubFPintarGridNumeracionCorporativa
            Else
                MsgBox "Error al consultar la numeración corporativa.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Consultar la información de la numeración pública
            If Me.proDatosProducto.MetConsultarDatosProductoNumero Then
                Call SubFPintarGridNumeroPublico
            Else
                MsgBox "Error al consultar la numeración corporativa.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Consultar del Plan de Numeración Actual
            Set varPlanNumeracion = New colPlanNumeracion
            Set varPlanNumeracion.proConexion = Me.proConexion
            varPlanNumeracion.proCliente = Me.proOnyx.ContactID
            
            If varPlanNumeracion.MetConsultaActuales Then
                Call SubFPintarTreePlanNumeracion
            Else
                MsgBox "Error al consultar el plan de numeración actual del cliente.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Consulta de servicios suplementarios
            If Me.proDatosProducto.proServiciosxNumero Is Nothing Then
                Set Me.proDatosProducto.proServiciosxNumero = New colServiciosxNumero
                Set Me.proDatosProducto.proServiciosxNumero.proConexion = Me.proConexion
                
                If Me.proDatosProducto.MetConsultarServiciosxNumero Then
                    Call SubFPintarGridServiciosxNumero
                    'Invoca la consulta de los servicios suplementarios
                    Set Me.proServiciosSuplementarios = New colServiciosSup
                    Set Me.proServiciosSuplementarios.proConexion = Me.proConexion
                    If Me.proServiciosSuplementarios.FunGConsultaTodos Then
                        Call SubFPintarGridServiciosSuplementarios
                    End If
                Else
                    MsgBox "Error al consultar los servicios suplementarios de cada número.", vbCritical, App.Title
                    Exit Sub
                End If
            Else
                Call SubFPintarGridServiciosxNumero
            End If
        Else
            MsgBox "Error al consultar la información del encabezado.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarTreePlanNumeracion()
    Dim Nodo As Node
    Dim varContador As Integer
    Dim varEnlaceAnterior As String
    Dim varVirtual As String 'Agregado por Carlos Castelblanco 2006/07/26
    On Error GoTo ErrManager
    
    Me.trvPlanActual.Nodes.Clear
    varEnlaceAnterior = ""
    For varContador = 1 To varPlanNumeracion.Count
        'If-Else-End If Agregado por Carlos Castelblanco 2006/07/26
        If varPlanNumeracion.Item(varContador).proVirtual = "S" Then
            varVirtual = " - Virtual"
        Else
            varVirtual = " - No Virtual"
        End If
                
        If varContador = 1 Then
            varEnlaceAnterior = varPlanNumeracion.Item(varContador).proSerialNumber
            
            Set Nodo = Me.trvPlanActual.Nodes.Add(, , varEnlaceAnterior, "[" & varEnlaceAnterior & "] - [" & varPlanNumeracion.Item(varContador).proAlias & "]")
            Set Nodo = Me.trvPlanActual.Nodes.Add(varEnlaceAnterior, tvwChild, varEnlaceAnterior & varPlanNumeracion.Item(varContador).proMarcacion, varPlanNumeracion.Item(varContador).proMarcacion & varVirtual)
        Else
            If varEnlaceAnterior = varPlanNumeracion.Item(varContador).proSerialNumber Then
                Set Nodo = Me.trvPlanActual.Nodes.Add(varEnlaceAnterior, tvwChild, varEnlaceAnterior & varPlanNumeracion.Item(varContador).proMarcacion, varPlanNumeracion.Item(varContador).proMarcacion & varVirtual)
            Else
                varEnlaceAnterior = varPlanNumeracion.Item(varContador).proSerialNumber
                Set Nodo = Me.trvPlanActual.Nodes.Add(, , varEnlaceAnterior, "[" & varEnlaceAnterior & "] - [" & varPlanNumeracion.Item(varContador).proAlias & "]")
                Set Nodo = Me.trvPlanActual.Nodes.Add(varEnlaceAnterior, tvwChild, varEnlaceAnterior & varPlanNumeracion.Item(varContador).proMarcacion, varPlanNumeracion.Item(varContador).proMarcacion & varVirtual)
            End If
        End If
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridNumeracionCorporativa()
    On Error GoTo ErrManager
    
    With Me.grdNumeracionPrivada
        .Cols = 3
        .Rows = 1
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 1000
        .TextMatrix(0, 0) = "Datos ProductoId"
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 2000
        .TextMatrix(0, 1) = "Marcarción"
        
        'Columna agregada por Carlos Castelblanco 2006/07/26:
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = "Virtual"
        
        .Col = 0
    End With
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub SubFInicializarGridIncidentes()
    On Error GoTo ErrManager
    
    With Me.grdAsuntosModificaron
        .Cols = 5
        .Rows = 1
        
        'Columna 1
        .Row = 0
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 1515
        .TextMatrix(0, 0) = "ID Asunto"
    
        'Columna 2
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 4695
        .TextMatrix(0, 1) = "Descripcion Asunto"
        
        'Columna 3
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1815
        .TextMatrix(0, 2) = "Tipo de Asunto"
    
        'Columna 4
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 2550
        .TextMatrix(0, 3) = "Categoria del Asunto"
    
        'Columna 5
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 2805
        .TextMatrix(0, 4) = "Fecha Modificación"
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridNumeroPublico()
    On Error GoTo ErrManager
    
    With Me.grdNumeracionPublica
        .Rows = 1
        .Cols = 6
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 0
        .TextMatrix(0, 0) = "DatosProductoId"
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = "Codigo Ciudad"
        
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1100
        .TextMatrix(0, 2) = "Ciudad"
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 1200
        .TextMatrix(0, 3) = "Número"
        
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 2050
        .TextMatrix(0, 4) = "Clasificación"
        
        .Col = 5
        .CellAlignment = 4
        .ColWidth(5) = 2050
        .TextMatrix(0, 5) = "Fecha Asignación"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridNumeroPublico()
    Dim varContador As Integer
    Dim varDescripcionClasificaciones As String
    On Error GoTo ErrManager
    
    Me.grdNumeracionPublica.Rows = 1
    
    For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
        varDescripcionClasificaciones = ""
        If Not IsNull(Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proClasificacionDescripcion) Then
            'varDescripcionClasificaciones =
        End If
        Me.grdNumeracionPublica.AddItem Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proDatosProductoId & vbTab & _
                                            Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionCode & vbTab & _
                                            Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionName & vbTab & _
                                            Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proNumero & vbTab & _
                                            Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proClasificacionDescripcion & vbTab & _
                                            Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proFechaAsignacion
                                    
        'Pierre Torres Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proClasificacionDescripcion
    Next varContador
    
    Me.grdNumeracionPublica.Row = 0
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarGridServiciosSuplementarios()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdNomenServiciosSuple.Rows = 1
    
    For varContador = 1 To Me.proServiciosSuplementarios.Count
        Me.grdNomenServiciosSuple.AddItem Me.proServiciosSuplementarios.Item(varContador).proiServicioSuplementarioId & vbTab & _
                                            Me.proServiciosSuplementarios.Item(varContador).provchNombreServicio
                                    
    Next varContador
    
    Me.grdNomenServiciosSuple.Row = 0
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridServiciosxNumero()
    Dim varContador As Integer
    Dim varRegionCodeAnterior As String
    Dim varRegionNameAnterior As String
    Dim varNumeroAnterior As String
    Dim varServiciosRelacionados As String
    
    On Error GoTo ErrManager
    
        For varContador = 1 To Me.proDatosProducto.proServiciosxNumero.Count
            If varContador = 1 Then
                varRegionCodeAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionCode
                varNumeroAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNumero
                varRegionNameAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionName
                'varServiciosRelacionados = "[" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNombreServicio & "],"
                varServiciosRelacionados = "[" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proServicioSuplementarioId & "],"
                ' fnp  20060913
                If Me.proDatosProducto.proServiciosxNumero.Count = 1 Then
                    Me.grdServiciosSuplementarios.AddItem Me.proDatosProducto.proServiciosxNumero.Item(varContador).proDatosProductoId & vbTab & _
                                          Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionCode & vbTab & _
                                          Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionName & vbTab & _
                                          Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNumero & vbTab & _
                                          Mid(varServiciosRelacionados, 1, Len(varServiciosRelacionados) - 1)
                End If
                
            Else
                If varRegionCodeAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionCode _
                   And varNumeroAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNumero Then
                   
                    'varServiciosRelacionados = varServiciosRelacionados & " [" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNombreServicio & "],"
                    varServiciosRelacionados = varServiciosRelacionados & " [" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proServicioSuplementarioId & "],"
                   
                    If varContador = Me.proDatosProducto.proServiciosxNumero.Count Then
                        Me.grdServiciosSuplementarios.AddItem Me.proDatosProducto.proServiciosxNumero.Item(varContador).proDatosProductoId & vbTab & _
                                              Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionCode & vbTab & _
                                              Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionName & vbTab & _
                                              Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNumero & vbTab & _
                                              Mid(varServiciosRelacionados, 1, Len(varServiciosRelacionados) - 1)
                    End If
                Else
                    Me.grdServiciosSuplementarios.AddItem Me.proDatosProducto.proServiciosxNumero.Item(varContador).proDatosProductoId & vbTab & _
                                                          varRegionCodeAnterior & vbTab & _
                                                          varRegionNameAnterior & vbTab & _
                                                          varNumeroAnterior & vbTab & _
                                                          Mid(varServiciosRelacionados, 1, Len(varServiciosRelacionados) - 1)
                    
                    varRegionCodeAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionCode
                    varNumeroAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNumero
                    varRegionNameAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionName
                    
                    'varServiciosRelacionados = "[" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNombreServicio & "],"
                    varServiciosRelacionados = "[" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proServicioSuplementarioId & "],"
                    
                    If varContador = Me.proDatosProducto.proServiciosxNumero.Count Then
                        Me.grdServiciosSuplementarios.AddItem Me.proDatosProducto.proServiciosxNumero.Item(varContador).proDatosProductoId & vbTab & _
                                              Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionCode & vbTab & _
                                              Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionName & vbTab & _
                                              Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNumero & vbTab & _
                                              Mid(varServiciosRelacionados, 1, Len(varServiciosRelacionados) - 1)
                    End If
                End If
            End If
        Next varContador
    
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridServiciosSuplementarios()
    On Error GoTo ErrManager
    
    With Me.grdServiciosSuplementarios
        .Rows = 1
        .Cols = 5
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 0
        .TextMatrix(0, 0) = "DatosProductoId"
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = "Codigo Ciudad"
        
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = "Ciudad"
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 1100
        .TextMatrix(0, 3) = "Número"
        
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 2000
        .TextMatrix(0, 4) = "Servicios Asignados"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridNomenServiciosSuplementarios()
    On Error GoTo ErrManager
    
    With Me.grdNomenServiciosSuple
        .Rows = 1
        .Cols = 2
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = "ID"
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 3000
        .TextMatrix(0, 1) = "Servicio"
        
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridNumeracionCorporativa()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdNumeracionPrivada.Rows = 1
    For varContador = 1 To Me.proDatosProducto.proNumeracionCorporativa.Count
        Me.grdNumeracionPrivada.AddItem Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proDatosProductoId & vbTab & _
                                        Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proMarcacion & vbTab & _
                                        Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proVirtual
                                        'Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proVirtual Agregado por Carlos Castelblanco 2006/07/26

    Next varContador
    
    Me.grdNumeracionPrivada.Col = 0
    Me.grdNumeracionPrivada.Row = 0
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridIncidentes()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdAsuntosModificaron.Rows = 1
    
    For varContador = 1 To Me.proDatosProducto.proDatosProductoIncident.Count
        Me.grdAsuntosModificaron.AddItem Me.proDatosProducto.proDatosProductoIncident.Item(varContador).proIncidentId & vbTab & _
                                         Me.proDatosProducto.proDatosProductoIncident.Item(varContador).proDescripcion & vbTab & _
                                         Me.proDatosProducto.proDatosProductoIncident.Item(varContador).proTipo & vbTab & _
                                         Me.proDatosProducto.proDatosProductoIncident.Item(varContador).proCategoria & vbTab & _
                                         Me.proDatosProducto.proDatosProductoIncident.Item(varContador).proFechaModificacion
    Next varContador
    
    Me.grdAsuntosModificaron.Row = 0
    'If Me.grdAsuntosModificaron.Rows >= 2 Then
    '    Call SubFPintarFila(Me.grdAsuntosModificaron, 1, Me.lblPintar.BackColor)
    'End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub SubFPintarGridDetalles()
    Dim varContador As Integer
    Dim varValor As String
    Dim varValorLista As EDCAdminVoz.claValor
    Dim varContadorAux As Integer
    Dim varValorCampo As String
    On Error GoTo ErrManager
    
    If Me.proDatosProducto.proParametrosProducto Is Nothing Then
        Exit Sub
    End If
    
    If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
        Exit Sub
    End If
    
    If Me.proDatosProducto.proDetalleDatosProducto Is Nothing Then
        Exit Sub
    End If
    
    varValor = ""
    Me.grdDatosVoz.Redraw = False
    Me.grdDatosVoz.Rows = 1
    For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
        
        varValor = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId & vbTab & _
                   Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId

        For varContadorAux = 1 To Me.proDatosProducto.proParametrosProducto.Count
            Select Case Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proTipo
                Case "L"
                    Set varValorLista = New EDCAdminVoz.claValor
                    Set varValorLista.proConexion = Me.proConexion
                
                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proCampo)
                        Case "vchUser1"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1
                        Case "vchUser2"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2
                        Case "vchUser3"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3
                        Case "vchUser4"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4
                        Case "vchUser5"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5
                        Case "vchUser6"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6
                        Case "vchUser7"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7
                        Case "vchUser8"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8
                        Case "vchUser9"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9
                        Case "vchUser10"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10
                        Case "vchUser11"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11
                        Case "vchUser12"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12
                        Case "vchUser13"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13
                        Case "vchUser14"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14
                        Case "vchUser15"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15
                        Case "vchUser16"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16
                        Case "vchUser17"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17
                        Case "vchUser18"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18
                        Case "vchUser19"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19
                        Case "vchUser20"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20
                        Case "vchUser21"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21
                        Case "vchUser22"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22
                        Case "vchUser23"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23
                        Case "vchUser24"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24
                        Case "vchUser25"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25
                        Case "vchUser26"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26
                        Case "vchUser27"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27
                        Case "vchUser28"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28
                        Case "vchUser29"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29
                        Case "vchUser30"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30
                        Case "vchUser31"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31
                        Case "vchUser32"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32
                        Case "vchUser33"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33
                        Case "vchUser34"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34
                        Case "vchUser35"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35
                        Case "vchUser36"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36
                        Case "vchUser37"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37
                        Case "vchUser38"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38
                        Case "vchUser39"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39
                        Case "vchUser40"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40
                    End Select
                    
                    If varValorLista.MetConsultar Then
                        varValor = varValor & vbTab & varValorLista.proValorDesc
                        Set varValorLista = Nothing
                    Else
                        MsgBox "Error al consultar el valor.", vbCritical, App.Title
                        Exit Sub
                    End If
                    
                Case "B"
                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proCampo)
                        Case "vchUser1"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser2"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser3"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser4"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser5"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser6"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser7"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser8"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser9"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser10"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser11"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser12"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser13"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser14"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser15"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser16"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser17"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser18"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser19"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser20"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser21"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser22"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser23"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser24"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser25"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser26"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser27"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser28"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser29"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser30"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser31"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser32"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser33"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser34"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser35"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser36"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser37"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser38"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser39"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser40"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                    End Select
                
                    varValor = varValor & vbTab & varValorCampo
                Case Else
                
                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proCampo)
                        Case "vchUser1"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1
                        Case "vchUser2"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2
                        Case "vchUser3"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3
                        Case "vchUser4"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4
                        Case "vchUser5"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5
                        Case "vchUser6"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6
                        Case "vchUser7"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7
                        Case "vchUser8"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8
                        Case "vchUser9"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9
                        Case "vchUser10"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10
                        Case "vchUser11"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11
                        Case "vchUser12"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12
                        Case "vchUser13"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13
                        Case "vchUser14"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14
                        Case "vchUser15"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15
                        Case "vchUser16"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16
                        Case "vchUser17"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17
                        Case "vchUser18"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18
                        Case "vchUser19"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19
                        Case "vchUser20"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20
                        Case "vchUser21"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21
                        Case "vchUser22"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22
                        Case "vchUser23"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23
                        Case "vchUser24"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24
                        Case "vchUser25"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25
                        Case "vchUser26"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26
                        Case "vchUser27"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27
                        Case "vchUser28"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28
                        Case "vchUser29"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29
                        Case "vchUser30"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30
                        Case "vchUser31"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31
                        Case "vchUser32"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32
                        Case "vchUser33"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33
                        Case "vchUser34"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34
                        Case "vchUser35"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35
                        Case "vchUser36"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36
                        Case "vchUser37"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37
                        Case "vchUser38"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38
                        Case "vchUser39"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39
                        Case "vchUser40"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40
                    End Select
                
                    varValor = varValor & vbTab & varValorCampo
            End Select
        Next varContadorAux
        
        Me.grdDatosVoz.AddItem varValor
        
        'No mostrar los regsitros eliminados
        If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proRecordStatus = 0 Then
            Me.grdDatosVoz.RowHeight(Me.grdDatosVoz.Rows - 1) = 0
        End If
        
        'No mostrar los registros cancelados
        If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId = "C" And Me.chkCancelados.Value = False Then
            Me.grdDatosVoz.RowHeight(Me.grdDatosVoz.Rows - 1) = 0
        Else
            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId = "C" Then
                Call SubFPintarFila(Me.grdDatosVoz, Me.grdDatosVoz.Rows - 1, Me.lblCancelados.BackColor)
            End If
        End If
    Next
    
    Me.grdDatosVoz.Redraw = True
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridDetalle()
    Dim varContador As Integer
    On Error GoTo ErrManager
        
        If Me.proDatosProducto.proParametrosProducto Is Nothing Then
            Exit Sub
        End If
        
        If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
            MsgBox "El producto del incidente seleccionado no tiene campos parametrizados."
            Exit Sub
        Else
       
            With Me.grdDatosVoz
                .Cols = Me.proDatosProducto.proParametrosProducto.Count + 2
                .Rows = 1
                .Row = 0
                
                .Col = 0
                .CellAlignment = 4
                .ColWidth(0) = 0
                .TextMatrix(0, 0) = "Código"
                
                .Col = 1
                .CellAlignment = 4
                .ColWidth(1) = 0
                .TextMatrix(0, 1) = "Codigo Estado"
                
                For varContador = 1 To Me.proDatosProducto.proParametrosProducto.Count
                    .Col = varContador
                    .CellAlignment = 4
                    If Me.proDatosProducto.proParametrosProducto.Item(varContador).proTipo = "F" Then
                        .ColWidth(varContador + 1) = 2000
                    Else
                        .ColWidth(varContador + 1) = 1500
                    End If
                    .TextMatrix(0, varContador + 1) = Me.proDatosProducto.proParametrosProducto.Item(varContador).proEtiqueta
                Next varContador
                
                .Col = 0
                .Row = 0
            End With
        End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Sub SubFSeguridad()
Dim varCuenta As Integer
Dim varEncontro As Boolean
On Error GoTo ErrorManager

    'Por default oculta el botón de administración
    Me.cmdAdministracion.Visible = False
    
    'instancia la colección de usuarios
    Set varColUsuarios = New EDCAdminVoz.colUsuario
    Set varColUsuarios.proConexion = Me.proConexion
    varColUsuarios.proAplicacionId = APlicacionId
    
    'Revisa los usuarios autorizados
    If varColUsuarios.FunGConsultaxApp Then
        varCuenta = 1
        varEncontro = False
        varGAdministracion = False
        varGValidacion = False
        While varCuenta <= varColUsuarios.Count And varEncontro = False
            If Trim(varColUsuarios(varCuenta).proUserId) = Trim(Me.proOnyx.UserLogin) Then
                'Busca los permisos
                varColUsuarios(varCuenta).proPrivilegios = Trim(varColUsuarios(varCuenta).proPrivilegios)
                varGAdministracion = (Left(varColUsuarios(varCuenta).proPrivilegios, 1) = "1")
                varGValidacion = (Mid(varColUsuarios(varCuenta).proPrivilegios, 3, 1) = "1")
                varEncontro = True
            Else
                varCuenta = varCuenta + 1
            End If
        Wend
        
        'Permisos para ver el botón de administración
        If varGAdministracion Then Me.cmdAdministracion.Visible = True
        
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Sub SubFAprobacionNumeros()
Dim varCuenta As Integer
Dim varEncontro As Boolean
On Error GoTo ErrorManager

    'Por default oculta el botón de administración
    Me.cmdAprobacionNumeros.Visible = False
    
    'instancia la colección de usuarios
    Set varColUsuarios = New EDCAdminVoz.colUsuario
    Set varColUsuarios.proConexion = Me.proConexion
    varColUsuarios.proAplicacionId = APlicacionId
    
    'Revisa los usuarios autorizados para aprobar numeros por clasificacion
    If varColUsuarios.FunGConsultaAprobacionNumeros Then
        varCuenta = 1
        varEncontro = False
        varGApruebaNumerosClasificacion = False
        While varCuenta <= varColUsuarios.Count And varEncontro = False
            If Trim(varColUsuarios(varCuenta).proUserId) = Trim(Me.proOnyx.UserLogin) Then
                varEncontro = True
                varGApruebaNumerosClasificacion = True
            Else
                varCuenta = varCuenta + 1
            End If
        Wend
        
        'Permisos para ver el botón de administración
        If varGApruebaNumerosClasificacion Then Me.cmdAprobacionNumeros.Visible = True
        
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrManager
    
    Set Me.proDatosProducto = Nothing
    'Set Me.proConexion = Nothing
    'Set Me.proOnyx = Nothing
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdAsuntosModificaron_Click()
    On Error GoTo ErrManager
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub subFActivarTLinea()
On Error GoTo ErrorManager

    'Cambia los colores por los colores vivos
    Me.pnlTLinea.BackColor = Me.pnlVerde.BackColor
    Me.lblTLinea.BackColor = Me.pnlVerde.BackColor
    
    'ENvia arriba la imagen viva
    Me.imgTlinea.ZOrder 0
    
    'Pone los controles arriba
    Me.frmTLinea.ZOrder 0
    Me.pnlCorporativa.ZOrder 0
    Me.pnlPublica.ZOrder 0
    Me.pnlTLinea.ZOrder 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub subFDesactivarTLinea()
On Error GoTo ErrorManager

    'Cambia los colores por los colores vivos
    Me.pnlTLinea.BackColor = Me.pnlGris.BackColor
    Me.lblTLinea.BackColor = Me.pnlGris.BackColor
    
    'ENvia arriba la imagen viva
    Me.imgTLineaDark.ZOrder 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub subFActivarPublica()
On Error GoTo ErrorManager

    'Cambia los colores por los colores vivos
    Me.pnlPublica.BackColor = Me.pnlAmarillo.BackColor
    Me.lblPublica.BackColor = Me.pnlAmarillo.BackColor
    
    'ENvia arriba la imagen viva
    Me.imgPublica.ZOrder 0
    
    'Pone los controles arriba
    Me.frmPublica.ZOrder 0
    Me.pnlTLinea.ZOrder 0
    Me.pnlCorporativa.ZOrder 0
    Me.pnlPublica.ZOrder 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
Private Sub subFDesactivarPublica()
On Error GoTo ErrorManager

    'Cambia los colores por los colores vivos
    Me.pnlPublica.BackColor = Me.pnlGris.BackColor
    Me.lblPublica.BackColor = Me.pnlGris.BackColor
    
    'ENvia arriba la imagen viva
    Me.imgPublicaDark.ZOrder 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub subFActivarCorporativa()
On Error GoTo ErrorManager

    'Cambia los colores por los colores vivos
    Me.pnlCorporativa.BackColor = Me.pnlRosa.BackColor
    Me.lblCorporativa.BackColor = Me.pnlRosa.BackColor
    
    'ENvia arriba la imagen viva
    Me.imgCorporativa.ZOrder 0
    
    'Pone los controles arriba
    Me.frmCorporativa.ZOrder 0
    Me.pnlPublica.ZOrder 0
    Me.pnlTLinea.ZOrder 0
    Me.pnlCorporativa.ZOrder 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub subFDesactivarCorporativa()
On Error GoTo ErrorManager

    'Cambia los colores por los colores vivos
    Me.pnlCorporativa.BackColor = Me.pnlGris.BackColor
    Me.lblCorporativa.BackColor = Me.pnlGris.BackColor
    
    'ENvia arriba la imagen viva
    Me.imgCorporativaDark.ZOrder 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


Private Sub imgCorporativaDark_Click()
On Error GoTo ErrorManager

   
    'Desactiva los demas
    subFDesactivarPublica
    subFDesactivarTLinea
    'Activa la corporativa
    subFActivarCorporativa
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub imgPublicaDark_Click()
On Error GoTo ErrorManager

   
    'Desactiva los demas
    subFDesactivarCorporativa
    subFDesactivarTLinea
    'Activa la corporativa
    subFActivarPublica
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


Private Sub imgTLineaDark_Click()
On Error GoTo ErrorManager

   
    'Desactiva los demas
    subFDesactivarPublica
    subFDesactivarCorporativa
    
    'Activa la corporativa
    subFActivarTLinea
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

