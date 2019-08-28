VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetalleDatosProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAB TELEFONIA"
   ClientHeight    =   11265
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   15045
   Icon            =   "frmDetalleDatosProducto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11265
   ScaleWidth      =   15045
   Begin VB.Frame fraTituloEncabezado 
      BackColor       =   &H00C09258&
      Caption         =   "  Información  del TAB  "
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
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   15195
   End
   Begin VB.Frame fraFondoEncabezado 
      Height          =   1515
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   14955
      Begin VB.ComboBox cboCodigoUso 
         Height          =   315
         ItemData        =   "frmDetalleDatosProducto.frx":0CCA
         Left            =   2175
         List            =   "frmDetalleDatosProducto.frx":0CD1
         Style           =   2  'Dropdown List
         TabIndex        =   218
         Top             =   510
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.ComboBox cboCodigoEstracto 
         Height          =   315
         ItemData        =   "frmDetalleDatosProducto.frx":0CE0
         Left            =   2190
         List            =   "frmDetalleDatosProducto.frx":0CE7
         Style           =   2  'Dropdown List
         TabIndex        =   217
         Top             =   150
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.ComboBox cboEstratos 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmDetalleDatosProducto.frx":0CF6
         Left            =   750
         List            =   "frmDetalleDatosProducto.frx":0CFD
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   2595
      End
      Begin VB.ComboBox cboUso 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmDetalleDatosProducto.frx":0D0C
         Left            =   750
         List            =   "frmDetalleDatosProducto.frx":0D13
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   490
         Width           =   2595
      End
      Begin Threed.SSPanel SSLocal 
         Height          =   1005
         Left            =   11760
         TabIndex        =   202
         Top             =   120
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   1773
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
         BevelOuter      =   0
         Begin VB.TextBox TxtEnlace 
            Height          =   285
            Left            =   960
            MaxLength       =   7
            TabIndex        =   12
            Top             =   690
            Width           =   945
         End
         Begin VB.TextBox TxtIdVenta 
            Height          =   285
            Left            =   960
            MaxLength       =   8
            TabIndex        =   11
            Top             =   405
            Width           =   945
         End
         Begin VB.CommandButton CmdModificarLocal 
            Caption         =   "&Modificar"
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
            Left            =   2160
            TabIndex        =   13
            Top             =   705
            Width           =   900
         End
         Begin VB.Label lblTItulo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "id Venta :"
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
            Left            =   120
            TabIndex        =   205
            Top             =   465
            Width           =   765
         End
         Begin VB.Label lblTItulo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "Enlace :"
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
            Left            =   120
            TabIndex        =   204
            Top             =   735
            Width           =   765
         End
         Begin VB.Label lblTItulo 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
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
            Index           =   14
            Left            =   120
            TabIndex        =   203
            Top             =   0
            Width           =   2880
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1035
         Left            =   3360
         TabIndex        =   17
         Top             =   90
         Width           =   8355
         _Version        =   65536
         _ExtentX        =   14737
         _ExtentY        =   1826
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
         BorderWidth     =   1
         BevelInner      =   1
         Begin Threed.SSPanel pnlCliente 
            Height          =   1140
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   10395
            _Version        =   65536
            _ExtentX        =   18336
            _ExtentY        =   2011
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
            BevelOuter      =   0
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
               TabIndex        =   33
               Top             =   1500
               Visible         =   0   'False
               Width           =   1185
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
               TabIndex        =   32
               Top             =   1770
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.CommandButton cmdEliminarCliente 
               Caption         =   "Elimi&nar"
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
               Left            =   7320
               TabIndex        =   31
               Top             =   720
               Width           =   900
            End
            Begin VB.CommandButton cmdBuscarCliente 
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
               Left            =   7320
               TabIndex        =   30
               Top             =   480
               Width           =   900
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
               TabIndex        =   48
               Top             =   1830
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
               TabIndex        =   47
               Top             =   1800
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
               TabIndex        =   46
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
               TabIndex        =   45
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
               TabIndex        =   44
               Top             =   1560
               Visible         =   0   'False
               Width           =   885
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
               TabIndex        =   43
               Top             =   1530
               Visible         =   0   'False
               Width           =   1545
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
               Left            =   4800
               TabIndex        =   42
               Top             =   720
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
               Left            =   5640
               TabIndex        =   9
               Top             =   720
               Width           =   1545
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
               Left            =   4800
               TabIndex        =   41
               Top             =   480
               Width           =   765
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
               Left            =   5640
               TabIndex        =   7
               Top             =   450
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
               Left            =   2280
               TabIndex        =   40
               Top             =   480
               Width           =   765
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
               Left            =   3120
               TabIndex        =   6
               Top             =   450
               Width           =   1545
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
               TabIndex        =   39
               Top             =   1080
               Visible         =   0   'False
               Width           =   9000
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
               TabIndex        =   38
               Top             =   1290
               Visible         =   0   'False
               Width           =   6255
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
               TabIndex        =   37
               Top             =   1530
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
               TabIndex        =   36
               Top             =   1500
               Visible         =   0   'False
               Width           =   1635
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
               TabIndex        =   35
               Top             =   1800
               Visible         =   0   'False
               Width           =   975
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
               TabIndex        =   34
               Top             =   1800
               Visible         =   0   'False
               Width           =   4050
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
               Left            =   975
               TabIndex        =   8
               Top             =   720
               Width           =   3690
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
               Left            =   150
               TabIndex        =   29
               Top             =   750
               Width           =   765
            End
            Begin VB.Label lblIDCliente 
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
               Left            =   960
               TabIndex        =   5
               Top             =   450
               Width           =   1185
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
               Left            =   150
               TabIndex        =   28
               Top             =   480
               Width           =   765
            End
            Begin VB.Label lblTItulo 
               BackColor       =   &H00C8D0D4&
               Caption         =   "Indique aquí el cliente que tiene o tendrá registrado el producto de telefonía nacional"
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
               Index           =   1
               Left            =   150
               TabIndex        =   27
               Top             =   240
               Width           =   6255
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
               Left            =   120
               TabIndex        =   26
               Top             =   15
               Width           =   8145
            End
         End
         Begin VB.CommandButton cmdGuardarEnvio 
            Caption         =   "&Guardar"
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdCancelarEnvio 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   4080
            TabIndex        =   23
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdModificarEnvio 
            Caption         =   "&Modificar"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkEnvioCorpLD 
            Caption         =   "Envío Corporativo LD"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3060
            TabIndex        =   21
            Top             =   360
            Width           =   2085
         End
         Begin VB.CheckBox chkEnvioCorpLocal 
            Caption         =   "Envío Corporativo Local"
            Enabled         =   0   'False
            Height          =   225
            Left            =   3060
            TabIndex        =   20
            Top             =   120
            Width           =   2265
         End
         Begin VB.CheckBox chkEnvioPublicoLD 
            Caption         =   "Envío Público LD"
            Enabled         =   0   'False
            Height          =   195
            Left            =   570
            TabIndex        =   19
            Top             =   420
            Width           =   1665
         End
         Begin VB.CheckBox chkEnvioPublicoLocal 
            Caption         =   "Envío Público Local"
            Enabled         =   0   'False
            Height          =   225
            Left            =   570
            TabIndex        =   18
            Top             =   120
            Width           =   2025
         End
      End
      Begin VB.TextBox txtComentarios 
         Enabled         =   0   'False
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
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1148
         Width           =   2235
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
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
         Left            =   1110
         TabIndex        =   3
         Top             =   830
         Width           =   1275
      End
      Begin VB.Label lblCiudadInstalacion 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4890
         TabIndex        =   10
         Top             =   1178
         Width           =   3150
      End
      Begin VB.Label lblTItulo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Ciudad Instalación :"
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
         Height          =   240
         Index           =   17
         Left            =   3390
         TabIndex        =   220
         Top             =   1185
         Width           =   1455
      End
      Begin VB.Label lblTItulo 
         BackColor       =   &H00808080&
         Caption         =   "Uso"
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
         Height          =   240
         Index           =   19
         Left            =   90
         TabIndex        =   219
         Top             =   535
         Width           =   630
      End
      Begin VB.Label lblTItulo 
         BackColor       =   &H00808080&
         Caption         =   "Estrato"
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
         Height          =   240
         Index           =   18
         Left            =   90
         TabIndex        =   216
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lblComentarios 
         AutoSize        =   -1  'True
         BackColor       =   &H00B98D1C&
         Caption         =   "Comentarios"
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
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   1185
         Width           =   975
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00B98D1C&
         Caption         =   "Código"
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
         Height          =   210
         Left            =   90
         TabIndex        =   15
         Top             =   870
         Width           =   975
      End
   End
   Begin TabDlg.SSTab TbFondo 
      Height          =   9585
      Left            =   120
      TabIndex        =   49
      Top             =   1650
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16907
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "TIPOS DE LINEA"
      TabPicture(0)   =   "frmDetalleDatosProducto.frx":0D22
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblSeccion(0)"
      Tab(0).Control(1)=   "fraFondoModificacion"
      Tab(0).Control(2)=   "fraFondoProductos(0)"
      Tab(0).Control(3)=   "pnlFuerte"
      Tab(0).Control(4)=   "pnlTenue"
      Tab(0).Control(5)=   "SSPanel2(0)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "NUMERACION PRIVADA"
      TabPicture(1)   =   "frmDetalleDatosProducto.frx":0D3E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSeccion(1)"
      Tab(1).Control(1)=   "fraFondoProductos(1)"
      Tab(1).Control(2)=   "Frame2(0)"
      Tab(1).Control(3)=   "SSPanel2(1)"
      Tab(1).Control(4)=   "SSPanel4"
      Tab(1).Control(5)=   "cmdRefrescarPlanesNumeracion"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "NUMERACION PUBLICA"
      TabPicture(2)   =   "frmDetalleDatosProducto.frx":0D5A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblSeccion(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraFondoProductos(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "SSPanel5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame2(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "SSPanel2(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmDetalleDatosProducto.frx":0D76
         Left            =   600
         List            =   "frmDetalleDatosProducto.frx":0D7D
         Style           =   2  'Dropdown List
         TabIndex        =   215
         Top             =   -1080
         Width           =   2280
      End
      Begin VB.CommandButton cmdRefrescarPlanesNumeracion 
         Caption         =   "&Refrescar Planes de Numeración"
         Height          =   285
         Left            =   -74130
         TabIndex        =   201
         Top             =   8760
         Width           =   2925
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Index           =   2
         Left            =   5940
         TabIndex        =   187
         Top             =   4440
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   767
         _StockProps     =   15
         BackColor       =   12620376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtIncidente 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   2460
            TabIndex        =   188
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lblIncidente 
            BackStyle       =   0  'Transparent
            Caption         =   "Incidente que está modificando:"
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
            Index           =   2
            Left            =   120
            TabIndex        =   189
            Top             =   90
            Width           =   2385
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C09258&
         Height          =   5145
         Index           =   1
         Left            =   60
         TabIndex        =   165
         Top             =   4350
         Width           =   9405
         Begin VB.Frame fraFondoProducto 
            Height          =   525
            Index           =   2
            Left            =   -30
            TabIndex        =   179
            Top             =   -90
            Width           =   9435
            Begin VB.TextBox txtNombreProducto 
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   1950
               TabIndex        =   183
               Top             =   150
               Width           =   3825
            End
            Begin VB.TextBox txtCodigoProducto 
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   780
               TabIndex        =   181
               Top             =   150
               Width           =   1155
            End
            Begin VB.Label lblProducto 
               AutoSize        =   -1  'True
               Caption         =   "Producto:"
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
               Index           =   2
               Left            =   60
               TabIndex        =   184
               Top             =   180
               Width           =   690
            End
         End
         Begin VB.Frame fraBotonesModificacion 
            Height          =   4275
            Index           =   2
            Left            =   6030
            TabIndex        =   166
            Top             =   780
            Width           =   2865
            Begin VB.CommandButton cmdPublicar 
               Caption         =   "&Publicar"
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
               Left            =   60
               TabIndex        =   178
               Top             =   2175
               Width           =   2715
            End
            Begin VB.CommandButton cmdInsertar 
               Caption         =   "&Agregar No Público"
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   60
               TabIndex        =   180
               Top             =   2460
               Width           =   2715
            End
            Begin VB.CommandButton cmdSeleccionarTodosModificacion 
               Caption         =   "Se&leccionar Todos "
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
               Index           =   2
               Left            =   90
               TabIndex        =   176
               Top             =   150
               Width           =   2685
            End
            Begin VB.Frame Frame1 
               Height          =   1365
               Index           =   2
               Left            =   180
               TabIndex        =   167
               Top             =   750
               Width           =   2535
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Líneas seleccionadas"
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
                  Index           =   2
                  Left            =   360
                  TabIndex        =   175
                  Top             =   1050
                  Width           =   1575
               End
               Begin VB.Label lblSeleccionModificacion 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   2
                  Left            =   90
                  TabIndex        =   174
                  Top             =   1080
                  Width           =   165
               End
               Begin VB.Label lblEliminar 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   2
                  Left            =   90
                  TabIndex        =   173
                  Top             =   810
                  Width           =   165
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Líneas a eliminar"
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
                  Index           =   2
                  Left            =   360
                  TabIndex        =   172
                  Top             =   780
                  Width           =   1200
               End
               Begin VB.Label lblModificar 
                  BackColor       =   &H00F9FCE7&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   2
                  Left            =   90
                  TabIndex        =   171
                  Top             =   510
                  Width           =   165
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Líneas en modificación"
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
                  Index           =   2
                  Left            =   360
                  TabIndex        =   170
                  Top             =   480
                  Width           =   1650
               End
               Begin VB.Label lblInsertar 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   2
                  Left            =   90
                  TabIndex        =   169
                  Top             =   210
                  Width           =   165
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Líneas agregadas"
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
                  Index           =   2
                  Left            =   360
                  TabIndex        =   168
                  Top             =   180
                  Width           =   1305
               End
            End
            Begin VB.CommandButton cmdDeseleccionarTodosModificacion 
               Caption         =   "De&seleccionar Todos"
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
               Index           =   2
               Left            =   90
               TabIndex        =   177
               Top             =   420
               Width           =   2685
            End
            Begin VB.CommandButton cmdDeshacerModificación 
               Caption         =   "Des&hacer  Modificaciones"
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   60
               TabIndex        =   182
               Top             =   3630
               Width           =   2715
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdEdicionNumeroPublico 
            Height          =   4275
            Left            =   150
            TabIndex        =   185
            Top             =   780
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   7541
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "NUMERACION PRIVADA EN EDICION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   186
            Top             =   540
            Width           =   2985
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   8805
         Left            =   9720
         TabIndex        =   162
         Top             =   690
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   15531
         _StockProps     =   15
         Caption         =   "SSPanel4"
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
         BevelInner      =   1
         Begin VB.CommandButton cmdModificarInsertados 
            Caption         =   "Modi&ficar Servicios Suplementarios"
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   2220
            TabIndex        =   198
            Top             =   8430
            Width           =   2715
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   645
            Index           =   4
            Left            =   2520
            TabIndex        =   163
            Top             =   30
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   1138
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
            BevelInner      =   1
            Begin VB.Image Image1 
               Height          =   480
               Index           =   5
               Left            =   90
               Picture         =   "frmDetalleDatosProducto.frx":0D8C
               Stretch         =   -1  'True
               Top             =   90
               Width           =   2235
            End
         End
         Begin Threed.SSPanel pnlExplicacion 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   190
            Top             =   420
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
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
            Height          =   7695
            Left            =   330
            TabIndex        =   164
            Top             =   720
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   13573
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
         End
      End
      Begin VB.Frame fraFondoProductos 
         BackColor       =   &H00808080&
         Height          =   3945
         Index           =   2
         Left            =   90
         TabIndex        =   145
         Top             =   360
         Width           =   14835
         Begin Threed.SSPanel pnlPublica 
            Height          =   1485
            Left            =   270
            TabIndex        =   196
            Top             =   720
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
            Begin VB.Image imgPublica 
               Height          =   1110
               Left            =   90
               MouseIcon       =   "frmDetalleDatosProducto.frx":3548
               Picture         =   "frmDetalleDatosProducto.frx":3E12
               Stretch         =   -1  'True
               Top             =   90
               Width           =   1365
            End
            Begin VB.Label lblPublica 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D8F7F8&
               Caption         =   "TPBCL         "
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
               TabIndex        =   197
               Top             =   1200
               Width           =   1365
            End
         End
         Begin VB.Frame FraTituloModificados 
            BackColor       =   &H00808080&
            Caption         =   "  Tipos de Linea actualmente en el cliente  "
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
            Index           =   2
            Left            =   -30
            TabIndex        =   146
            Top             =   0
            Width           =   14895
         End
         Begin Threed.SSPanel pnlExplicacion 
            Height          =   375
            Index           =   3
            Left            =   1080
            TabIndex        =   147
            Top             =   330
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
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
         Begin MSFlexGridLib.MSFlexGrid grdNumeroPublico 
            Height          =   3255
            Left            =   1770
            TabIndex        =   148
            Top             =   630
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   5741
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
         End
         Begin VB.Frame fraBotones 
            Height          =   3285
            Index           =   2
            Left            =   6540
            TabIndex        =   149
            Top             =   630
            Width           =   3015
            Begin VB.CommandButton CmdCambiarTipoLinea 
               Caption         =   "Cambiar Tipo &Linea"
               Height          =   285
               Left            =   60
               TabIndex        =   221
               Top             =   1920
               Width           =   2115
            End
            Begin VB.CommandButton cmdSeleccionarTodos 
               Caption         =   "Seleccionar &Todos "
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
               Index           =   2
               Left            =   60
               TabIndex        =   160
               Top             =   150
               Width           =   2145
            End
            Begin VB.CommandButton cmdDeseleccionarTodos 
               Caption         =   "Deseleccionar T&odos"
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
               Index           =   2
               Left            =   60
               TabIndex        =   161
               Top             =   420
               Width           =   2145
            End
            Begin VB.Frame Frame3 
               Height          =   795
               Index           =   2
               Left            =   60
               TabIndex        =   154
               Top             =   720
               Width           =   2715
               Begin VB.Label lblSinSeleccion 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   2
                  Left            =   3600
                  TabIndex        =   159
                  Top             =   180
                  Width           =   165
               End
               Begin VB.Label lblSeleccion 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   2
                  Left            =   90
                  TabIndex        =   158
                  Top             =   480
                  Width           =   165
               End
               Begin VB.Label Label4 
                  Caption         =   "Líneas seleccionadas"
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
                  Index           =   2
                  Left            =   300
                  TabIndex        =   157
                  Top             =   450
                  Width           =   2385
               End
               Begin VB.Label Label2 
                  Caption         =   "Líneas canceladas"
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
                  Index           =   2
                  Left            =   300
                  TabIndex        =   156
                  Top             =   180
                  Width           =   2385
               End
               Begin VB.Label lblCancelados 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   2
                  Left            =   90
                  TabIndex        =   155
                  Top             =   210
                  Width           =   165
               End
            End
            Begin VB.CommandButton cmdModificarColumna 
               Caption         =   "Mo&dificar Columna"
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   9090
               TabIndex        =   153
               Top             =   480
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.CommandButton cmdClonar 
               Caption         =   "&Clonar Registros"
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   9090
               TabIndex        =   152
               Top             =   150
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.CommandButton cmdEliminar 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   60
               TabIndex        =   151
               Top             =   1620
               Width           =   2115
            End
            Begin VB.CommandButton cmdModificar 
               Caption         =   "&Modificar"
               Enabled         =   0   'False
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
               Index           =   2
               Left            =   9330
               TabIndex        =   150
               Top             =   330
               Visible         =   0   'False
               Width           =   1245
            End
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   8055
         Left            =   -74910
         TabIndex        =   143
         Top             =   660
         Width           =   4485
         _Version        =   65536
         _ExtentX        =   7911
         _ExtentY        =   14208
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
         BevelInner      =   1
         Begin Threed.SSPanel pnlExplicacion 
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   144
            Top             =   240
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
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
         Begin Threed.SSPanel pnlExplicacion 
            Height          =   375
            Index           =   5
            Left            =   450
            TabIndex        =   191
            Top             =   3960
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Plan de numeración en proceso de instalación"
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
            Height          =   3345
            Left            =   150
            TabIndex        =   199
            Top             =   540
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   5900
            _Version        =   393217
            HideSelection   =   0   'False
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            BorderStyle     =   1
            Appearance      =   1
         End
         Begin MSComctlLib.TreeView trvPlanEnCurso 
            Height          =   3645
            Left            =   150
            TabIndex        =   200
            Top             =   4260
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   6429
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Index           =   1
         Left            =   -64440
         TabIndex        =   140
         Top             =   4440
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   767
         _StockProps     =   15
         BackColor       =   12620376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtIncidente 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   2460
            TabIndex        =   141
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lblIncidente 
            BackStyle       =   0  'Transparent
            Caption         =   "Incidente que está modificando:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   142
            Top             =   90
            Width           =   2385
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C09258&
         Height          =   5145
         Index           =   0
         Left            =   -70350
         TabIndex        =   119
         Top             =   4350
         Width           =   10215
         Begin VB.Frame fraBotonesModificacion 
            Height          =   4275
            Index           =   1
            Left            =   7170
            TabIndex        =   125
            Top             =   780
            Width           =   2865
            Begin VB.CommandButton cmdDeshacerModificación 
               Caption         =   "Des&hacer  Modificaciones"
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   60
               TabIndex        =   126
               Top             =   3630
               Width           =   2715
            End
            Begin VB.CommandButton cmdSeleccionarTodosModificacion 
               Caption         =   "Se&leccionar Todos "
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
               Index           =   1
               Left            =   90
               TabIndex        =   129
               Top             =   150
               Width           =   2685
            End
            Begin VB.CommandButton cmdDeseleccionarTodosModificacion 
               Caption         =   "De&seleccionar Todos"
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
               Index           =   1
               Left            =   90
               TabIndex        =   128
               Top             =   420
               Width           =   2685
            End
            Begin VB.Frame Frame1 
               Height          =   1365
               Index           =   1
               Left            =   180
               TabIndex        =   130
               Top             =   750
               Width           =   2535
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Extensiones agregadas"
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
                  Index           =   1
                  Left            =   360
                  TabIndex        =   138
                  Top             =   180
                  Width           =   1710
               End
               Begin VB.Label lblInsertar 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   1
                  Left            =   90
                  TabIndex        =   137
                  Top             =   210
                  Width           =   165
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Extensiones en modificación"
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
                  Index           =   1
                  Left            =   360
                  TabIndex        =   136
                  Top             =   480
                  Width           =   2055
               End
               Begin VB.Label lblModificar 
                  BackColor       =   &H00F9FCE7&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   1
                  Left            =   90
                  TabIndex        =   135
                  Top             =   510
                  Width           =   165
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Extensiones a eliminar"
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
                  Index           =   1
                  Left            =   360
                  TabIndex        =   134
                  Top             =   780
                  Width           =   1605
               End
               Begin VB.Label lblEliminar 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   1
                  Left            =   90
                  TabIndex        =   133
                  Top             =   810
                  Width           =   165
               End
               Begin VB.Label lblSeleccionModificacion 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   1
                  Left            =   90
                  TabIndex        =   132
                  Top             =   1080
                  Width           =   165
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Extensiones seleccionadas"
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
                  Index           =   1
                  Left            =   360
                  TabIndex        =   131
                  Top             =   1050
                  Width           =   1980
               End
            End
            Begin VB.CommandButton cmdInsertar 
               Caption         =   "&Agregar Extensión"
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   60
               TabIndex        =   127
               Top             =   2220
               Width           =   2715
            End
         End
         Begin VB.Frame fraFondoProducto 
            Height          =   525
            Index           =   1
            Left            =   0
            TabIndex        =   120
            Top             =   -90
            Width           =   10215
            Begin VB.TextBox txtCodigoProducto 
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   780
               TabIndex        =   122
               Top             =   150
               Width           =   1155
            End
            Begin VB.TextBox txtNombreProducto 
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   1950
               TabIndex        =   121
               Top             =   150
               Width           =   3825
            End
            Begin VB.Label lblProducto 
               AutoSize        =   -1  'True
               Caption         =   "Producto:"
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
               Index           =   1
               Left            =   60
               TabIndex        =   123
               Top             =   180
               Width           =   690
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdEdicionNumeracionPrivada 
            Height          =   4275
            Left            =   150
            TabIndex        =   124
            Top             =   780
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   7541
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "NUMERACION PRIVADA EN EDICION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   139
            Top             =   540
            Width           =   2985
         End
      End
      Begin VB.Frame fraFondoProductos 
         BackColor       =   &H00808080&
         Height          =   3945
         Index           =   1
         Left            =   -75000
         TabIndex        =   102
         Top             =   360
         Width           =   14835
         Begin Threed.SSPanel pnlExplicacion 
            Height          =   375
            Index           =   1
            Left            =   5580
            TabIndex        =   103
            Top             =   360
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
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
         Begin Threed.SSPanel pnlCorporativa 
            Height          =   1485
            Left            =   4710
            TabIndex        =   194
            Top             =   630
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
            Begin VB.Image imgCorporativa 
               Height          =   1110
               Left            =   90
               MouseIcon       =   "frmDetalleDatosProducto.frx":B4B0
               Picture         =   "frmDetalleDatosProducto.frx":BD7A
               Stretch         =   -1  'True
               Top             =   90
               Width           =   1305
            End
            Begin VB.Label lblCorporativa 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
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
               TabIndex        =   195
               Top             =   1200
               Width           =   1305
            End
         End
         Begin VB.Frame FraTituloModificados 
            BackColor       =   &H00808080&
            Caption         =   "  Extensiones instaladas en la oficina   "
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
            Index           =   0
            Left            =   -30
            TabIndex        =   104
            Top             =   0
            Width           =   14895
         End
         Begin MSFlexGridLib.MSFlexGrid grdNumeracionPrivada 
            Height          =   2475
            Left            =   6120
            TabIndex        =   105
            Top             =   660
            Width           =   8205
            _ExtentX        =   14473
            _ExtentY        =   4366
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
         End
         Begin VB.Frame fraBotones 
            Height          =   795
            Index           =   1
            Left            =   5790
            TabIndex        =   106
            Top             =   3060
            Width           =   8955
            Begin VB.CommandButton cmdModificar 
               Caption         =   "&Modificar"
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   9330
               TabIndex        =   118
               Top             =   330
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.CommandButton cmdEliminar 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   4620
               TabIndex        =   117
               Top             =   150
               Width           =   1395
            End
            Begin VB.CommandButton cmdClonar 
               Caption         =   "&Clonar Registros"
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   9090
               TabIndex        =   116
               Top             =   150
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.CommandButton cmdModificarColumna 
               Caption         =   "Mo&dificar Columna"
               Enabled         =   0   'False
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
               Index           =   1
               Left            =   9090
               TabIndex        =   115
               Top             =   480
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.Frame Frame3 
               Height          =   795
               Index           =   1
               Left            =   1860
               TabIndex        =   109
               Top             =   0
               Width           =   2715
               Begin VB.Label lblCancelados 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   1
                  Left            =   90
                  TabIndex        =   114
                  Top             =   210
                  Width           =   165
               End
               Begin VB.Label Label2 
                  Caption         =   "Extensiones canceladas"
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
                  Index           =   1
                  Left            =   300
                  TabIndex        =   113
                  Top             =   180
                  Width           =   2385
               End
               Begin VB.Label Label4 
                  Caption         =   "Extensiones seleccionadas"
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
                  Index           =   1
                  Left            =   300
                  TabIndex        =   112
                  Top             =   450
                  Width           =   2385
               End
               Begin VB.Label lblSeleccion 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   1
                  Left            =   90
                  TabIndex        =   111
                  Top             =   480
                  Width           =   165
               End
               Begin VB.Label lblSinSeleccion 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   1
                  Left            =   3600
                  TabIndex        =   110
                  Top             =   180
                  Width           =   165
               End
            End
            Begin VB.CommandButton cmdSeleccionarTodos 
               Caption         =   "Seleccionar &Todos "
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
               Index           =   1
               Left            =   60
               TabIndex        =   108
               Top             =   150
               Width           =   1785
            End
            Begin VB.CommandButton cmdDeseleccionarTodos 
               Caption         =   "Deseleccionar T&odos"
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
               Index           =   1
               Left            =   60
               TabIndex        =   107
               Top             =   450
               Width           =   1785
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Index           =   0
         Left            =   -64920
         TabIndex        =   96
         Top             =   4440
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   767
         _StockProps     =   15
         BackColor       =   12620376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtIncidente 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   2460
            TabIndex        =   97
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label lblIncidente 
            BackStyle       =   0  'Transparent
            Caption         =   "Incidente que está modificando:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   98
            Top             =   90
            Width           =   2385
         End
      End
      Begin Threed.SSPanel pnlTenue 
         Height          =   225
         Left            =   -68280
         TabIndex        =   58
         Top             =   9210
         Visible         =   0   'False
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   397
         _StockProps     =   15
         Caption         =   "SSPanel3"
         ForeColor       =   0
         BackColor       =   16644326
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
      Begin Threed.SSPanel pnlFuerte 
         Height          =   225
         Left            =   -68850
         TabIndex        =   57
         Top             =   9210
         Visible         =   0   'False
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   397
         _StockProps     =   15
         Caption         =   "SSPanel2"
         ForeColor       =   16777215
         BackColor       =   12620376
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
      Begin VB.Frame fraFondoProductos 
         BackColor       =   &H00808080&
         Height          =   3945
         Index           =   0
         Left            =   -74970
         TabIndex        =   50
         Top             =   360
         Width           =   14835
         Begin VB.Frame pnlGrupoCentrexCallSource 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            ForeColor       =   &H80000008&
            Height          =   1905
            Left            =   10350
            TabIndex        =   212
            Top             =   960
            Width           =   2685
            Begin VB.ListBox lstGrupoCentrexCallSource 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   1470
               Left            =   0
               TabIndex        =   214
               Top             =   450
               Width           =   2685
            End
            Begin VB.TextBox txtGrupoCentrexCallSource 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   0
               MultiLine       =   -1  'True
               TabIndex        =   213
               Top             =   0
               Width           =   2685
            End
         End
         Begin VB.CommandButton cmdLiberarRecursos 
            Caption         =   "&Liberar Recurso"
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
            Left            =   13140
            TabIndex        =   211
            Top             =   660
            Width           =   1635
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "&Guardar"
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
            Left            =   13140
            TabIndex        =   210
            Top             =   300
            Width           =   1635
         End
         Begin VB.TextBox txtCallSource 
            Height          =   285
            Left            =   11550
            TabIndex        =   209
            Top             =   600
            Width           =   1485
         End
         Begin VB.TextBox txtGrupoCentrex 
            Height          =   285
            Left            =   11550
            TabIndex        =   207
            Top             =   300
            Width           =   1485
         End
         Begin Threed.SSPanel pnlExplicacion 
            Height          =   405
            Index           =   0
            Left            =   570
            TabIndex        =   101
            Top             =   600
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "Los tipos de línea corresponden al tipo de conexión física entre la oficina y la red. Estos pueden ser E1, RDSI, Básicas etc."
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
         Begin Threed.SSPanel pnlTLinea 
            Height          =   1815
            Left            =   30
            TabIndex        =   192
            Top             =   960
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   3201
            _StockProps     =   15
            Caption         =   "Corporativa"
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
            Begin VB.Image imgTlinea 
               Height          =   1455
               Left            =   90
               MouseIcon       =   "frmDetalleDatosProducto.frx":C824
               Picture         =   "frmDetalleDatosProducto.frx":D0EE
               Top             =   270
               Width           =   1065
            End
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
               MousePointer    =   99  'Custom
               TabIndex        =   193
               Top             =   90
               Width           =   1065
            End
         End
         Begin VB.Frame FraTituloModificados 
            BackColor       =   &H00808080&
            Caption         =   "  Tipos de Linea actualmente en el cliente  "
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
            Index           =   1
            Left            =   -30
            TabIndex        =   59
            Top             =   0
            Width           =   14895
         End
         Begin MSFlexGridLib.MSFlexGrid grdDetalles 
            Height          =   2205
            Left            =   1230
            TabIndex        =   51
            Top             =   900
            Width           =   13575
            _ExtentX        =   23945
            _ExtentY        =   3889
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
         End
         Begin VB.Frame fraBotones 
            Height          =   795
            Index           =   0
            Left            =   1170
            TabIndex        =   60
            Top             =   3030
            Width           =   13545
            Begin VB.CommandButton cmdSeleccionarTodos 
               Caption         =   "Seleccionar &Todos "
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
               Index           =   0
               Left            =   60
               TabIndex        =   74
               Top             =   150
               Width           =   1785
            End
            Begin VB.CheckBox chkCancelados 
               Caption         =   "Ver líneas canceladas"
               Height          =   195
               Index           =   0
               Left            =   6270
               TabIndex        =   72
               Top             =   540
               Width           =   2025
            End
            Begin VB.Frame Frame3 
               Height          =   795
               Index           =   0
               Left            =   1920
               TabIndex        =   66
               Top             =   0
               Width           =   2865
               Begin VB.Label lblSinSeleccion 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   0
                  Left            =   3600
                  TabIndex        =   71
                  Top             =   180
                  Width           =   165
               End
               Begin VB.Label lblSeleccion 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   0
                  Left            =   150
                  TabIndex        =   70
                  Top             =   480
                  Width           =   165
               End
               Begin VB.Label Label4 
                  Caption         =   "Tipos de Líneas seleccionadas"
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
                  Index           =   0
                  Left            =   420
                  TabIndex        =   69
                  Top             =   450
                  Width           =   2385
               End
               Begin VB.Label Label2 
                  Caption         =   "Tipos de Líneas canceladas"
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
                  Index           =   0
                  Left            =   420
                  TabIndex        =   68
                  Top             =   180
                  Width           =   2385
               End
               Begin VB.Label lblCancelados 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   0
                  Left            =   150
                  TabIndex        =   67
                  Top             =   210
                  Width           =   165
               End
            End
            Begin VB.CommandButton cmdClonar 
               Caption         =   "&Clonar Registros"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   9090
               TabIndex        =   64
               Top             =   150
               Width           =   1605
            End
            Begin VB.TextBox txtCantidadRegistros 
               BackColor       =   &H00FDF8E6&
               Enabled         =   0   'False
               Height          =   285
               Left            =   8520
               TabIndex        =   63
               Top             =   180
               Width           =   465
            End
            Begin VB.CommandButton cmdEliminar 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   12060
               TabIndex        =   62
               Top             =   180
               Width           =   1395
            End
            Begin VB.CommandButton cmdModificar 
               Caption         =   "&Modificar"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   4860
               TabIndex        =   61
               Top             =   150
               Width           =   1245
            End
            Begin VB.CommandButton cmdModificarColumna 
               Caption         =   "Mo&dificar Columna"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   9090
               TabIndex        =   65
               Top             =   480
               Width           =   1605
            End
            Begin VB.CommandButton cmdDeseleccionarTodos 
               Caption         =   "Deseleccionar T&odos"
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
               Index           =   0
               Left            =   60
               TabIndex        =   73
               Top             =   420
               Width           =   1785
            End
            Begin VB.Label lblCantidadRegistros 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad de Registros Activos:"
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
               Index           =   0
               Left            =   6270
               TabIndex        =   75
               Top             =   180
               Width           =   2235
            End
         End
         Begin VB.Label lblCallSource 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Call Source:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   10440
            TabIndex        =   208
            Top             =   660
            Width           =   1125
         End
         Begin VB.Label lblGrupoCentrex 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Grupo Centrex:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   10260
            TabIndex        =   206
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Frame fraFondoModificacion 
         BackColor       =   &H00C09258&
         Height          =   5235
         Left            =   -74940
         TabIndex        =   52
         Top             =   4320
         Width           =   14775
         Begin VB.Frame fraFondoProducto 
            Height          =   525
            Index           =   0
            Left            =   -30
            TabIndex        =   92
            Top             =   -90
            Width           =   14895
            Begin VB.TextBox txtNombreProducto 
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   2820
               TabIndex        =   94
               Top             =   150
               Width           =   5775
            End
            Begin VB.TextBox txtCodigoProducto 
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   1080
               TabIndex        =   93
               Top             =   150
               Width           =   1725
            End
            Begin VB.Label lblProducto 
               AutoSize        =   -1  'True
               Caption         =   "Producto:"
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
               Index           =   0
               Left            =   210
               TabIndex        =   95
               Top             =   180
               Width           =   690
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdDetallesModificacion 
            Height          =   3795
            Left            =   120
            TabIndex        =   53
            Top             =   660
            Width           =   14565
            _ExtentX        =   25691
            _ExtentY        =   6694
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
         End
         Begin VB.Frame fraBotonesModificacion 
            Height          =   825
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   4320
            Width           =   14565
            Begin VB.CommandButton cmdInsertar 
               Caption         =   "&Agregar Tipos de Línea"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   7290
               TabIndex        =   100
               Top             =   150
               Width           =   1875
            End
            Begin VB.CommandButton cmdModificarInsertados 
               Caption         =   "Modi&ficar"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   9270
               TabIndex        =   90
               Top             =   150
               Width           =   2265
            End
            Begin VB.CommandButton cmdSeleccionarTodosModificacion 
               Caption         =   "Se&leccionar Todos "
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
               Index           =   0
               Left            =   90
               TabIndex        =   88
               Top             =   150
               Width           =   1785
            End
            Begin VB.CommandButton cmdClonarModificados 
               Caption         =   "Clonar Re&gistros"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   11730
               TabIndex        =   87
               Top             =   150
               Width           =   2655
            End
            Begin VB.Frame Frame1 
               Height          =   825
               Index           =   0
               Left            =   1920
               TabIndex        =   78
               Top             =   0
               Width           =   5325
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipos de Línea seleccionadas"
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
                  Index           =   0
                  Left            =   3060
                  TabIndex        =   86
                  Top             =   480
                  Width           =   2145
               End
               Begin VB.Label lblSeleccionModificacion 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   0
                  Left            =   2790
                  TabIndex        =   85
                  Top             =   510
                  Width           =   165
               End
               Begin VB.Label lblEliminar 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   0
                  Left            =   2790
                  TabIndex        =   84
                  Top             =   210
                  Width           =   165
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipos de Línea a eliminar"
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
                  Index           =   0
                  Left            =   3060
                  TabIndex        =   83
                  Top             =   180
                  Width           =   1770
               End
               Begin VB.Label lblModificar 
                  BackColor       =   &H00F9FCE7&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   0
                  Left            =   90
                  TabIndex        =   82
                  Top             =   510
                  Width           =   165
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipos de Línea en modificación"
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
                  Index           =   0
                  Left            =   390
                  TabIndex        =   81
                  Top             =   480
                  Width           =   2220
               End
               Begin VB.Label lblInsertar 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   165
                  Index           =   0
                  Left            =   90
                  TabIndex        =   80
                  Top             =   210
                  Width           =   165
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipos de Línea agregadas"
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
                  Index           =   0
                  Left            =   360
                  TabIndex        =   79
                  Top             =   180
                  Width           =   1875
               End
            End
            Begin VB.CommandButton cmdModificarColumnaInsertados 
               Caption         =   "Mo&dificar Columna"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   9270
               TabIndex        =   91
               Top             =   420
               Width           =   2265
            End
            Begin VB.CommandButton cmdDeshacerModificación 
               Caption         =   "Des&hacer  Modificaciones"
               Enabled         =   0   'False
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
               Index           =   0
               Left            =   11730
               TabIndex        =   77
               Top             =   420
               Width           =   2655
            End
            Begin VB.CommandButton cmdDeseleccionarTodosModificacion 
               Caption         =   "De&seleccionar Todos"
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
               Index           =   0
               Left            =   90
               TabIndex        =   89
               Top             =   420
               Width           =   1785
            End
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "TIPOS DE LINEA EN EDICION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   99
            Top             =   450
            Width           =   2985
         End
      End
      Begin VB.Label lblSeccion 
         Alignment       =   2  'Center
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERACION PUBLICA"
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
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   56
         Top             =   30
         Width           =   1875
      End
      Begin VB.Label lblSeccion 
         Alignment       =   2  'Center
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERACION PRIVADA"
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
         Height          =   285
         Index           =   1
         Left            =   -73590
         TabIndex        =   55
         Top             =   30
         Width           =   1935
      End
      Begin VB.Label lblSeccion 
         Alignment       =   2  'Center
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIPOS DE LINEA"
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
         Height          =   285
         Index           =   0
         Left            =   -74970
         TabIndex        =   54
         Top             =   30
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmDetalleDatosProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'   DESCRIPCION         : En este formulario se configuran los diferentes tipos de linea del clie-
'                         nte y se asocian los numeros publicos y extenciones en caso de corporat-
'                         ivo, se tienen do conceptos que difiden el fomulario en 2 partes en toda
'                         su funcionalidad: en curso(tipos de linea, extenciones y numeros publicos
'                         asigandos al cliete pero sinconfiguracion fisica) y en uso( recursos de
'                         vos asignados y configurados ya en el cliente), cuando se terminan los pro
'                         cesos tenminan en curso actualiza el Uso.
'   PARAMETROS          :
'                        proDatosProducto       ClaDatosProducto
'                        proConexion            ClaConexion
'                        proOnyx                ClaONYX
'                        proClienteTelefonia    ClaClienteTelefonia
'
'   RETORNO             : NA
'
'   EJEMPLO             :
'                        Set frmDetalleDatosProducto.proDatosProducto = Me.proDatosProducto
'                        Set frmDetalleDatosProducto.proConexion = Me.proConexion
'                        Set frmDetalleDatosProducto.proOnyx = Me.proOnyx
'                        Set frmDetalleDatosProducto.proClienteTelefonia = Me.proClienteTelefonia
'                        frmDetalleDatosProducto.Show (1)

'*************************************************************************************************
'   MODIFICADO POR      : Carlos Leonardo Villamil (I&T)
'   DESCRIPCION CAMBIO  : El cambio permite cambiar la asociacion entre tipos de linea en curso o
'                         en uso y numeros publicos en uso.
'   VERCION             : 3.7.4
'   FECHA               : 09-JUL-09
'*************************************************************************************************

'**********************************************************************
' MODIFICADO POR :      CARLOS ALBERTO BARRERA
' DESCRIPCION CAMBIO:   frmEdicionDetalleDatos.proiClienteId = Me.lblIDCliente
' VERSION: 1.0.000
' FECHA: JUNIO 30/2009
'****************************************************************


'**********************************************************************
' MODIFICADO POR :      IVAN MAURICIO FONSECA I&T
' DESCRIPCION CAMBIO:   Con el cambio se controla que solo los tipos con la categoria de incidentes
'                       correctos puedan modificar la numeración.
' VERSION: 2.0.000
' FECHA: 15/Junio/2012
'****************************************************************
Option Explicit

Public proConexion As ADODB.Connection
Public proDatosProducto As claDatosProducto
Public proOnyx As EDCVoz.claONYX
Public proClienteTelefonia As EDCVoz.claClienteTelefonia
Public proEstrato As EDCAdminVoz.colEstratoCiudad
Public proUsoServicio As EDCAdminVoz.colEstratos

Private varColClienteTelefonia As EDCVoz.colClienteTelefonia
Private varValoresCampoProductoTipoLinea As EDCAdminVoz.colValoresCampoProducto

Private varNumeros As EDCAdminVoz.colNumero
Private varConsultaNumeros As EDCAdminVoz.claConsultaNumero
Private varPlanNumeracion As colPlanNumeracion
Private varPlanNumeracionEnCurso As colPlanNumeracionEnCurso
Private varClasificacion As EDCAdminVoz.colClasificacion

Private varOperacionOnyx As EDCAdminVoz.colOperacionOnyx
Private varInsertarTiposLinea As Boolean
Private varModificarTiposLinea As Boolean
Private varEliminarTiposLinea As Boolean
Private varInsertarNumeracionCorporativa As Boolean
Private varModificarNumeracionCorporativa As Boolean
Private varEliminarNumeracionCorporativa As Boolean
Private varInsertarNumeracionPublica As Boolean
Private varModificarNumeracionPublica As Boolean
Private varEliminarNumeracionPublica As Boolean
Private varTiposLinea As Boolean
Private varNumerosPrivados As Boolean
Private varNumerosPublicos As Boolean
Private varSi As String

Private Sub cboEstratos_Click()
On Error GoTo ErrManager
    Me.cboCodigoEstracto.ListIndex = Me.cboEstratos.ListIndex
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cboUso_Click()
On Error GoTo ErrManager
    Me.cboCodigoUso.ListIndex = Me.cboUso.ListIndex
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub chkCancelados_Click(Index As Integer)
    On Error GoTo ErrManager
    
    If Index = 0 Then
        Call SubFPintarGridTiposLinea(Index)
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub chkEnvioCorpLD_Click()
    On Error GoTo ErrManager
    
    Me.cmdGuardarEnvio.Enabled = True
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub chkEnvioCorpLocal_Click()
    On Error GoTo ErrManager
    
    Me.cmdGuardarEnvio.Enabled = True
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub chkEnvioPublicoLD_Click()
    On Error GoTo ErrManager
    
    Me.cmdGuardarEnvio.Enabled = True
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub chkEnvioPublicoLocal_Click()
    On Error GoTo ErrManager
    
    Me.cmdGuardarEnvio.Enabled = True
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdBuscarCliente_Click()
Dim varCliente As EDCTraslados.claCliente
Dim varcontenidos As Integer
On Error GoTo ErrorManager

    'Validar campos requeridos
    If cboEstratos.ListIndex < 0 Or cboUso.ListIndex < 0 Then
        MsgBox "Debe configurar el estrato y uso de servicio.", vbCritical, App.Title
        Exit Sub
    End If
    
    Set varCliente = Nothing
    Set varCliente = New EDCTraslados.claCliente
    Set varCliente.proConexion = Me.proConexion
    
    Set frmBuscarCliente.proCliente = varCliente
    frmBuscarCliente.Show vbModal
    
    If Val(varCliente.proClienteId) = 0 Then Exit Sub
    
    Me.proDatosProducto.proClienteNacionalId = varCliente.proClienteId
    Me.proDatosProducto.proNombreClienteNacional = varCliente.proNombreCliente
    
    'Actualizar en la clase valores modificados
    Me.proDatosProducto.proiEstratoid = cboCodigoEstracto.Text
    Me.proDatosProducto.proUsoServicioId = cboCodigoUso.Text
    
    'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
    If Not Me.proDatosProducto.MetGuardar Then
        MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Inserta o actualiza la información de los incidentes
    If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
        MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
        Exit Sub
    End If
    
     MsgBox "Se agregó al cliente " & Me.proDatosProducto.proNombreClienteNacional & " como vínculo para la Telefonía Nacional.", vbInformation, App.Title
            Me.lblIDCliente = Me.proDatosProducto.proClienteNacionalId
            Me.lblCliente = Me.proDatosProducto.proNombreClienteNacional
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdBuscarClienteLocal_Click()
Dim varCliente As EDCTraslados.claCliente
Dim varcontenidos As Integer
On Error GoTo ErrorManager

    Set varCliente = Nothing
    Set varCliente = New EDCTraslados.claCliente
    Set varCliente.proConexion = Me.proConexion
    
    Set frmBuscarCliente.proCliente = varCliente
    frmBuscarCliente.Show vbModal
    
    If Val(varCliente.proClienteId) = 0 Then Exit Sub
    
    Me.proDatosProducto.proClienteLocalId = varCliente.proClienteId
    Me.proDatosProducto.proNombreClienteLocal = varCliente.proNombreCliente
    
    Me.lblIDClienteLocal.Caption = Me.proDatosProducto.proClienteLocalId
    Me.lblClienteLocal.Caption = Me.proDatosProducto.proNombreClienteLocal
    
    If Me.proDatosProducto.MetGuardar Then
            MsgBox "Se agregó al cliente " & Me.proDatosProducto.proNombreClienteLocal & " como vínculo para la Telefonía Local.", vbInformation, App.Title
            Me.lblIDCliente = Me.proDatosProducto.proClienteNacionalId
            Me.lblCliente = Me.proDatosProducto.proNombreClienteNacional
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


' -->Inicio 3.7.4

'************************************************************************************************
'   DESCRIPCION         : Este evento llama al formulario de seleccion de tipo de linea a donde
'                         cambiaran los numeros y asigan a los nuevos tipos de linea los numeros
'                         en curso.
'   PARAMETROS          :
'                         NA
'   RETORNO             :
'                         NA
'
'   EJEMPLO             :
'                        CmdCambiarTipoLinea_Click()

'*************************************************************************************************
'   MODIFICADO POR      : Carlos Leonardo Villamil (I&T)
'   DESCRIPCION CAMBIO  : El cambio permite cambiar la asociacion entre tipos de linea en curso o
'                         en uso y numeros publicos en uso.
'   VERCION             : 3.7.4
'   FECHA               : 09-JUL-09
'*************************************************************************************************
'*************************************************************************************************
'   MODIFICADO POR      : Carlos Leonardo Villamil (I&T)
'   DESCRIPCION CAMBIO  : Error en el declaracion  de variables se pasan a long
'   VERSION             : 1.0.101
'   FECHA               : 16-SEP-2009
'*************************************************************************************************



Private Sub CmdCambiarTipoLinea_Click()
    Dim varContador As Integer, varIndiceDetalleDatos As Integer, varIndiceTipoLinea As Integer, varIndice As Integer
    Dim varContadorAux As Integer
    Dim varEncontro As Boolean
    Dim varContadorSeleccionados As Integer
    Dim varNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
    Dim varNovedadNumeros As claNovedadNumero
    Dim varNovedadNumeracionCorporativa As claNovedadNumeracionCorporativa
    Dim varNuevoTipoLinea As Long
    Dim varBackup As Boolean
    Dim varRepuestaTiposLineafrm As Integer
    Dim varColValoresCampoProducto As EDCAdminVoz.colValoresCampoProducto
    Dim varTipoLineaSeleccionada As Long '-->1.0.101
    Dim varContadorTiposLinea As Long
    Dim varContadorNumeros As Long
    Dim varIndiceTipo As Long '-->1.0.101
    Dim varCantidadNumerosTiposdeLineaEnUso As Integer
    Dim varCantidadNumerosTiposdeLineaEnCurso As Integer
    Dim VarVchuser1 As String
    On Error GoTo ErrManager
        VarVchuser1 = "vchUser1"
        Set frmCambioTipoLinea.proDatosProducto = proDatosProducto
        frmCambioTipoLinea.Show vbModal
        varNuevoTipoLinea = frmCambioTipoLinea.proTipoLinea
        If varNuevoTipoLinea < 0 Then
             varRepuestaTiposLineafrm = MsgBox("No se eligio ningun tipo de linea, se cancela la transaccion", vbInformation, "Tipos Linea")
             Exit Sub
        End If
        
        'vlidar
        varTipoLineaSeleccionada = -1
        For varContadorTiposLinea = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            If varNuevoTipoLinea = Me.proDatosProducto.proDetalleDatosProducto.Item(varContadorTiposLinea).proDetalleDatosProductoId Then
                varTipoLineaSeleccionada = Me.proDatosProducto.proDetalleDatosProducto.Item(varContadorTiposLinea).proUser1
                        For varContadorNumeros = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
                            If Me.proDatosProducto.proDatosProductoNumero.Item(varContadorNumeros).proTipoLinea = varNuevoTipoLinea Then
                                varCantidadNumerosTiposdeLineaEnUso = varCantidadNumerosTiposdeLineaEnUso + 1
                            End If
                        Next varContadorNumeros
            End If
        Next varContadorTiposLinea
        If varTipoLineaSeleccionada = -1 Then
                For varContadorTiposLinea = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                    If varNuevoTipoLinea = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorTiposLinea).proNovedadDetalleDatosProductoId Then
                        varTipoLineaSeleccionada = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorTiposLinea).proUser1
                        varCantidadNumerosTiposdeLineaEnCurso = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorTiposLinea).proContadorNumeros
                    End If
                Next varContadorTiposLinea
        End If
        
        
        
        
        
        If varTipoLineaSeleccionada = -1 Then
             varRepuestaTiposLineafrm = MsgBox("Se selccionoun tipo de linea invalido , se cancela la transaccion", vbInformation, "Tipos Linea")
             Exit Sub
        End If
        Set varColValoresCampoProducto = New EDCAdminVoz.colValoresCampoProducto
        Set varColValoresCampoProducto.proConexion = Me.proConexion
        varColValoresCampoProducto.proCampo = VarVchuser1
        varColValoresCampoProducto.proValorIdPadre = 0
        varColValoresCampoProducto.proProductNumber = Me.proDatosProducto.proProductNumber
        varColValoresCampoProducto.MetConsultarValoresxProducto
        varColValoresCampoProducto.MetConsultarxCampoProducto
        varIndiceTipo = varColValoresCampoProducto.BuscarIndiceProValorId(Trim(Str(varTipoLineaSeleccionada)))
        If varColValoresCampoProducto.Item(varIndiceTipo).proMaximo < (varCantidadNumerosTiposdeLineaEnUso + varCantidadNumerosTiposdeLineaEnCurso + Me.proDatosProducto.proDatosProductoNumero.proSeleccionados) Then
                MsgBox "El tipo de línea no tiene la cantidad de números permitidos (" + Str(varColValoresCampoProducto.Item(varIndiceTipo).proMaximo) + " ), se cancela la transacción.", AccesoRestringido, "Tipos Linea"
                Exit Sub
        End If
        
        If Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los registros a eliminar.", vbInformation, App.Title
            Exit Sub
        End If
        If MsgBox("Desea cambiar de Tipo de linea los [" & Me.proDatosProducto.proDatosProductoNumero.proSeleccionados & "] detalles selecionados?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
            If Not Me.proDatosProducto.MetGuardar Then
                MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
                Exit Sub
            End If
            'Inserta o actualiza la información de los incidentes
            If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
                MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
            End If
            varContadorSeleccionados = 0
            For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
                If Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proSeleccion = "1" Then
                    varContadorSeleccionados = varContadorSeleccionados + 1
                    'Validar que ya no se encuentre dentro de los seleccionados para eliminar
                    varEncontro = False
                    For varContadorAux = 1 To Me.proDatosProducto.proNovedadNumero.Count
                        If Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionCode = _
                           Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proRegionCode _
                           And Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proNumero = _
                           Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proNumero _
                            Then
                            varEncontro = True
                            Exit For
                        End If
                    Next varContadorAux
                    If varEncontro = True Then
                        MsgBox "El [" & varContadorSeleccionados & "] registro seleccionado, ya se encuentra marcado para eliminación, modificación o cambio tipo de linea."
                    Else
                        Set varNovedadNumeros = Nothing
                        Set varNovedadNumeros = New claNovedadNumero
                        Set varNovedadNumeros.proConexion = Me.proConexion
                        varNovedadNumeros.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                        varNovedadNumeros.proFechaLiberacion = ""
                        varNovedadNumeros.proFechaReserva = ""
                        varNovedadNumeros.proIncidentId = Me.proDatosProducto.proIncidentId
                        varNovedadNumeros.proNovedadNumeroId = "0"
                        varNovedadNumeros.proNumero = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proNumero
                        varNovedadNumeros.proRegionCode = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionCode
                        varNovedadNumeros.proRegionName = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionName
                        varNovedadNumeros.proTipoNovedadId = 2 ' tipo novedad
                        varNovedadNumeros.proAsociaNovedad = "S"
                        varNovedadNumeros.proPublicar = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proPublicar
                        varNovedadNumeros.proTipoLineaAnterior = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea
                        'Revisar si se debe eliminar el tipo de línea del número eliminado (si no tiene más números asignados)
                        Dim varexiste As Boolean
                        varexiste = False
                        If Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea <> 0 Then 'Si tiene un tipo de línea asignado
                            varIndiceDetalleDatos = proDatosProducto.proDetalleDatosProducto.IndexOf(Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea)
                            If (proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proContador <= 1) Then
                                'Se debe revisar si el tipo de linea no se encuentra en edición
                                Dim varindicenovedad As Integer
                                varindicenovedad = 1
                                While varindicenovedad <= Me.proDatosProducto.proNovedadDetalleDatosProducto.Count And varexiste = False
                                    If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varindicenovedad).proDetalleDatosProductoId = proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proDetalleDatosProductoId Then
                                        varexiste = True
                                    End If
                                    varindicenovedad = varindicenovedad + 1
                                Wend
                                If varexiste = True Then
                                    MsgBox "No es posible modificar el tipo de linea de el número ''" & varNovedadNumeros.proNumero & "'' ya que el tipo de línea [" & Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea & "] se encuentra en edición.", vbOKOnly, App.Title
                                End If
                            End If
                        End If
                        If varexiste = False Then
                            varTiposLinea = False 'Consultar nuevamente pestaña de tipos de línea
                            If varIndiceDetalleDatos <> 0 Then
                                proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).FunGDisminuirContador
                            End If
                             varNovedadNumeros.proTipoLinea = varNuevoTipoLinea
                            If varNovedadNumeros.FunGInsertar Then
                                If Not Me.proDatosProducto.MetAgregarNovedadNumeroPublico(varNovedadNumeros) Then
                                    MsgBox "Error al agregar el detalle [" & varNovedadNumeros.proNumero & "].", vbCritical, App.Title
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Error al marcar el número [" & varNovedadNumeros.proNumero & "] para Cambiado de Tipo de linea.", vbInformation, App.Title
                            End If
                        End If
                    End If
                End If
            Next varContador
        End If
        For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
            Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proSeleccion = "0"
        Next varContador
    
        Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = 0
        
        Call SubFPintarGridNumeroPublico(2)
        Call SubFPintarGridEdicionNumeroPublico(2)
        
        Me.CmdCambiarTipoLinea.Enabled = False
        Me.cmdEliminar(2).Enabled = False

    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
' -->Fin 3.7.4



Private Sub cmdCancelarEnvio_Click()
    On Error GoTo ErrManager
    
    Me.chkEnvioCorpLD.Enabled = False
    Me.chkEnvioCorpLocal.Enabled = False
    Me.chkEnvioPublicoLD.Enabled = False
    Me.chkEnvioPublicoLocal.Enabled = False
    
    Me.cmdCancelarEnvio.Visible = False
    Me.cmdGuardarEnvio.Visible = False
    Me.cmdModificarEnvio.Visible = True
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdClonar_Click(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    'Tipos de líneas
    If Index = 0 Then
        If Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar el detalle que desea clonar.", vbInformation, App.Title
            Exit Sub
        End If
        
        If Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados > 1 Then
            MsgBox "Para ejecutar esta opción solo debe existir un registro seleccionado.", vbInformation, App.Title
            Exit Sub
        End If
        
        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "1" Then
                If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId <> "A" Then
                    MsgBox "Solo puede clonar registros activos.", vbInformation, App.Title
                    Exit Sub
                End If
                Exit For
            End If
        Next
        
        Set frmClonar.proConexion = Me.proConexion
        Set frmClonar.proDatosProducto = Me.proDatosProducto
        Set frmClonar.proOnyx = Me.proOnyx
        
        frmClonar.proOrigen = "O"
        frmClonar.proRegistro = varContador
        frmClonar.Show (vbModal)
        
        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0
        
        Call SubFPintarGridTiposLinea(Index)
        Call SubFPintarGridTiposLineaModificacion(Index)
        
        Me.cmdModificar(Index).Enabled = False
        Me.cmdModificarColumna(Index).Enabled = False
        Me.cmdEliminar(Index).Enabled = False
        Me.cmdClonar(Index).Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdClonarModificados_Click(Index As Integer)
    Dim RegistroSeleccionado As Integer
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    'Tipos de líneas
    If Index = 0 Then
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar el detalle que desea clonar.", vbInformation, App.Title
            Exit Sub
        End If
        
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados > 1 Then
            MsgBox "Para ejecutar esta opción solo debe existir un registro seleccionado.", vbInformation, App.Title
            Exit Sub
        End If
        
        For RegistroSeleccionado = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(RegistroSeleccionado).proSeleccion = "1" Then
                Exit For
            End If
        Next RegistroSeleccionado
        
        
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(RegistroSeleccionado).proStatusId <> "A" Then
            MsgBox "Solo puede clonar registros activos.", vbInformation, App.Title
            Exit Sub
        End If
        
        Set frmClonar.proConexion = Me.proConexion
        Set frmClonar.proDatosProducto = Me.proDatosProducto
        Set frmClonar.proOnyx = Me.proOnyx
        
        frmClonar.proOrigen = "M"
        frmClonar.proRegistro = RegistroSeleccionado
        frmClonar.Show (vbModal)
        
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0
        
        Call SubFPintarGridTiposLineaModificacion(Index)
        
        Me.cmdClonarModificados(Index).Enabled = False
        Me.cmdDeshacerModificación(Index).Enabled = False
        Me.cmdModificarInsertados(Index).Enabled = False
        Me.cmdModificarColumnaInsertados(Index).Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdDeseleccionarTodos_Click(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    'Tipos de líneas
    Select Case Index
        Case 0
            For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
                Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "0"
            Next varContador
            
            Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0
            
            Call SubFPintarGridTiposLinea(Index)
        Case 1
            For varContador = 1 To Me.proDatosProducto.proNumeracionCorporativa.Count
                Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proSeleccion = "0"
            Next varContador
            
            Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados = 0
            
            Call SubFPintarGridNumeracionCorporativa(Index)
        Case 2
            For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
                Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proSeleccion = "0"
            Next varContador
            
            Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = 0
            
            Call SubFPintarGridNumeroPublico(Index)
    End Select
    
    Me.cmdModificar(Index).Enabled = False
    Me.cmdModificarColumna(Index).Enabled = False
    Me.cmdEliminar(Index).Enabled = False
    Me.cmdClonar(Index).Enabled = False
    Me.CmdCambiarTipoLinea.Enabled = False '-->3.7.4

    
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub

Private Sub cmdDeseleccionarTodosModificacion_Click(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    'Tipos de líneas
    If Index = 0 Then
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0
        
        Call SubFPintarGridTiposLineaModificacion(Index)
        
        Me.cmdClonarModificados(Index).Enabled = False
        Me.cmdModificarColumnaInsertados(Index).Enabled = False
        Me.cmdModificarInsertados(Index).Enabled = False
    End If
    
    'Numeración Corporativa
    If Index = 1 Then
        For varContador = 1 To Me.proDatosProducto.proNovedadNumeracionCorporativa.Count
            Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados = 0
            
        Call SubFPintarGridEdicionNumeracionCorporativa(Index)
    End If
    
    'Numeración Pública
    If Index = 2 Then
        For varContador = 1 To Me.proDatosProducto.proNovedadNumero.Count
            Me.proDatosProducto.proNovedadNumero.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proNovedadNumero.proSeleccionados = 0
            
        Call SubFPintarGridEdicionNumeroPublico(Index)
    End If
    
    Me.cmdDeshacerModificación(Index).Enabled = False
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdDeshacerModificación_Click(Index As Integer)
    Dim varContador As Integer, varIndice As Integer, varIndiceDetalleDatosProducto As Integer
    Dim varFNumero As EDCAdminVoz.claNumero
    Dim varContinuar As Boolean
    Dim varDescripcionFAXTOMAIL As String
    Dim varDescripcionTELEFONOVIRTUAL As String
    Dim varNuevoEstado As String
    Dim varAlgunos As Boolean
    On Error GoTo ErrManager
    'Tipos de líneas
    varAlgunos = False
    If Index = 0 Then
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los registros a eliminar.", vbInformation, App.Title
            Exit Sub
        End If
        If MsgBox("Desea eliminar los [" & Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados & "] registros selecionados?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            varContador = 1
            Dim varNumero As String
            While varContador <= Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "1" Then
                   'Validacion que los numeros a eliminar no posean servicios suplementarios
                   Dim varvalido As Boolean
                   varvalido = True
                   If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proTipoNovedadId = "1" Then
                        varIndice = 1
                        If Me.proDatosProducto.proNovedadNumero Is Nothing Then Me.proDatosProducto.MetConsultarNovedadNumeros
                        If Not Me.proDatosProducto.proNovedadNumero Is Nothing Then
                            Do While varIndice <= Me.proDatosProducto.proNovedadNumero.Count And varvalido = True
                                If proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proNovedadDetalleDatosProductoId = proDatosProducto.proNovedadNumero.Item(varIndice).proTipoLinea Then
                                    If Me.proDatosProducto.proNovedadNumero.Item(varIndice).MetConsultarServiciosxReserva Then
                                        If Me.proDatosProducto.proNovedadNumero.Item(varIndice).procolServiciosxReserva.Count > 0 Then
                                            varvalido = False
                                            varNumero = Me.proDatosProducto.proNovedadNumero.Item(varIndice).proNumero
                                        End If
                                    Else
                                        MsgBox "Error al consultar  los servicios del número.", vbCritical, App.Title
                                        Exit Sub
                                    End If
                                End If
                                varIndice = varIndice + 1
                            Loop
                        End If
                    End If
                    If varvalido = False Then
                       MsgBox "No es posible realizar la acción: El Número [" & varNumero & "] no se puede eliminar, porque tiene servicios suplementarios.", vbInformation, App.Title
                       varContador = varContador + 1
                       varAlgunos = True
                    Else
                        If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).MetEliminar Then
                            Dim varindicedetalle As Integer
                            varindicedetalle = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId
                            If varindicedetalle <> 0 Then
                                Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proDatosProducto.proDetalleDatosProducto.IndexOf(varindicedetalle)).proEliminar = False
                                Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proDatosProducto.proDetalleDatosProducto.IndexOf(varindicedetalle)).proBackUp = False
                                Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proDatosProducto.proDetalleDatosProducto.IndexOf(varindicedetalle)).proModificar = False
                            End If
                            'Consultar números en proceso de instalación o modificación
                            If Me.proDatosProducto.proNovedadNumero Is Nothing Then
                                Set Me.proDatosProducto.proNovedadNumero = New colNovedadNumero
                                Set Me.proDatosProducto.proNovedadNumero.proConexion = Me.proConexion
                                
                                If Me.proDatosProducto.MetConsultarNovedadNumeros Then
                                    Call SubFPintarGridEdicionNumeroPublico(Index)
                                Else
                                    MsgBox "Error al consultar los números  que se encuentran en proceso de instalación o modificación.", vbCritical, App.Title
                                    Exit Sub
                                End If
                            End If
                            'Deshacer eliminación de números con el tipo de línea
                            varIndice = 1
                            Do While varIndice <= proDatosProducto.proNovedadNumero.Count
                                If proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proNovedadDetalleDatosProductoId = proDatosProducto.proNovedadNumero.Item(varIndice).proTipoLinea Or _
                                    ((proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId = proDatosProducto.proNovedadNumero.Item(varIndice).proTipoLineaAnterior) And proDatosProducto.proNovedadNumero.Item(varIndice).proTipoLineaAnterior <> 0) Then
                                    If proDatosProducto.proNovedadNumero.Item(varIndice).proTipoLineaAnterior <> 0 Then
                                        For varIndiceDetalleDatosProducto = 1 To proDatosProducto.proDetalleDatosProducto.Count
                                            If proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatosProducto).proDetalleDatosProductoId = proDatosProducto.proNovedadNumero.Item(varIndice).proTipoLineaAnterior Then
                                                proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatosProducto).FunGAumentarContador
                                                Exit For
                                            End If
                                        Next
                                    End If
                                    Set varFNumero = New claNumero
                                    Set varFNumero.proConexion = Me.proConexion
                                    varFNumero.proUpdateBy = Me.proOnyx.UserLogin
                                    varFNumero.proRecordStatus = 1
                                    If Me.proDatosProducto.proNovedadNumero.Item(varIndice).FunGEliminar Then
                                        'Liberar el número
                                        varFNumero.proUpdateDate = Format(Now, "mm/dd/yyyy hh:mm:ss")
                                        varFNumero.proRegionCode = Me.proDatosProducto.proNovedadNumero.Item(varIndice).proRegionCode
                                        varFNumero.proNumero = Me.proDatosProducto.proNovedadNumero.Item(varIndice).proNumero
                                        varNuevoEstado = "L"
                                        If varDescripcionFAXTOMAIL <> "" Then
                                            If InStr(1, Me.proDatosProducto.proNovedadNumero.Item(varIndice).proClasificacionDescripcion, varDescripcionFAXTOMAIL, vbTextCompare) > 0 Then
                                                varNuevoEstado = "F"
                                            End If
                                        End If
                                        If varDescripcionTELEFONOVIRTUAL <> "" Then
                                            If InStr(1, Me.proDatosProducto.proNovedadNumero.Item(varIndice).proClasificacionDescripcion, varDescripcionTELEFONOVIRTUAL, vbTextCompare) > 0 Then
                                                varNuevoEstado = "V"
                                            End If
                                        End If
                                        varFNumero.proEstadoNumero = varNuevoEstado
                                        If varFNumero.FunGModificar Then
                                            Me.proDatosProducto.proNovedadNumero.Remove (varIndice)
                                            Me.proDatosProducto.proNovedadNumero.proSeleccionados = Me.proDatosProducto.proNovedadNumero.proSeleccionados - 1
                                        Else
                                            MsgBox "Error al actualizar el estado del número.", vbCritical, App.Title
                                            Exit Do
                                        End If
                                    Else
                                        MsgBox "Error al eliminar el número [" + Me.proDatosProducto.proNovedadNumero.Item(varContador).proNumero + "].", vbCritical, App.Title
                                        Exit Do
                                    End If
                                Else
                                    varIndice = varIndice + 1
                                End If
                            Loop
                            
                            Me.proDatosProducto.proNovedadDetalleDatosProducto.Remove (varContador)
                            Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados - 1
                            
                        Else
                            MsgBox "Error al eliminar el registro.", vbCritical, App.Title
                            Exit Sub
                        End If
                    End If
                Else
                    varContador = varContador + 1
                End If
            Wend
            varNumerosPublicos = False 'Consultar nuevamente la grilla de números
            If varAlgunos = False Then
                MsgBox "Los registros se eliminaron exitosamente.", vbInformation, App.Title
            Else
                MsgBox "Proceso realizado exitosamente. Algunos registros no se pudieron eliminar.", vbInformation, App.Title
            End If
            For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "0"
            Next varContador
        
            Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0
        
            Call SubFPintarGridTiposLineaModificacion(Index)
            
            Me.cmdClonarModificados(Index).Enabled = False
            Me.cmdDeshacerModificación(Index).Enabled = False
            Me.cmdModificarInsertados(Index).Enabled = False
            Me.cmdModificarColumnaInsertados(Index).Enabled = False
        End If
    End If
    
    'Numeración Corporativa
    If Index = 1 Then
        If Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los números que desea eliminar.", vbInformation, App.Title
            Exit Sub
        End If
        
        
        If MsgBox("Desea eliminar los [" & Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados & "] registros selecionados?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            varContador = 1
            While varContador <= Me.proDatosProducto.proNovedadNumeracionCorporativa.Count
                If Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proSeleccion = "1" Then
                    If Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).FunGEliminar Then
                        Me.proDatosProducto.proNovedadNumeracionCorporativa.Remove (varContador)
                        Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados = Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados - 1
                    Else
                        MsgBox "Error al eliminar el número [" + Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proMarcacion + "].", vbCritical, App.Title
                    End If
                Else
                    varContador = varContador + 1
                End If
            Wend
            
            Call SubFPintarGridEdicionNumeracionCorporativa(Index)
            
            Me.cmdDeshacerModificación(Index).Enabled = False
        End If
    End If
    
    'Numeración Pública
    If Index = 2 Then
        If Me.proDatosProducto.proNovedadNumero.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los números que desea eliminar.", vbInformation, App.Title
            Exit Sub
        End If
        
        If MsgBox("Desea eliminar los [" & Me.proDatosProducto.proNovedadNumero.proSeleccionados & "] registros selecionados?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Set varClasificacion = New EDCAdminVoz.colClasificacion
            Set varClasificacion.proConexion = Me.proConexion
            If Not varClasificacion.FunGConsulta Then
                MsgBox "No fue posible consultar las clasificaciones.", vbError, App.Title
                Exit Sub
            End If
            varDescripcionFAXTOMAIL = ""
            varDescripcionTELEFONOVIRTUAL = "" ' 20070815 - CC
            For varContador = 1 To varClasificacion.Count
                If varClasificacion.Item(varContador).proClasificacionId = 3 Then
                    varDescripcionFAXTOMAIL = varClasificacion.Item(varContador).proClasificacion
                End If
                ' 20070815 - CC:
                If varClasificacion.Item(varContador).proClasificacionId = 4 Then
                    varDescripcionTELEFONOVIRTUAL = varClasificacion.Item(varContador).proClasificacion
                End If
                If varDescripcionFAXTOMAIL <> "" And varDescripcionTELEFONOVIRTUAL <> "" Then
                    Exit For
                End If
            Next
            Set varFNumero = New EDCAdminVoz.claNumero
            Set varFNumero.proConexion = Me.proConexion
            varFNumero.proUpdateBy = Me.proOnyx.UserLogin
            varFNumero.proRecordStatus = 1
            varContador = 1
            'Pierre Torres Me.proDatosProducto.proDatosProductoNumero.Item(0).proClasificacionDescripcion
            While varContador <= Me.proDatosProducto.proNovedadNumero.Count
                If Me.proDatosProducto.proNovedadNumero.Item(varContador).proSeleccion = "1" Then
                    varContinuar = True
                    If Me.proDatosProducto.proNovedadNumero.Item(varContador).MetConsultarServiciosxReserva Then
                        If Me.proDatosProducto.proNovedadNumero.Item(varContador).procolServiciosxReserva.Count > 0 Then
                            MsgBox "El Número [" & Me.proDatosProducto.proNovedadNumero.Item(varContador).proNumero & "] no se puede eliminar, porque tiene servicios suplementarios.", vbInformation, App.Title
                            varContinuar = False
                        End If
                    Else
                        MsgBox "Error al consultar  los servicios del número.", vbCritical, App.Title
                        Exit Sub
                    End If
                    If varContinuar Then
                        If Me.proDatosProducto.proNovedadNumero.Item(varContador).FunGEliminar Then
                            'Actualizar contadores de tipos de línea
                            If proDatosProducto.proNovedadNumero.Item(varContador).proTipoLineaAnterior = 0 Then
                                ' JM 11-Ene-2008 : Este FOR debe ser condicionado a si es un tipo de linea instalado o en novedad
                                If proDatosProducto.proNovedadNumero.Item(varContador).proAsociaNovedad = "S" Then
                                    For varIndice = 1 To proDatosProducto.proNovedadDetalleDatosProducto.Count
                                        If proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).proNovedadDetalleDatosProductoId = proDatosProducto.proNovedadNumero.Item(varContador).proTipoLinea Then
                                            proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).FunGDisminuirContador
                                            Exit For
                                        End If
                                    Next
                                Else
                                    ' JM 11-Ene-2008 : FOR que no existia y que debe ir
                                    For varIndice = 1 To proDatosProducto.proDetalleDatosProducto.Count
                                        If proDatosProducto.proDetalleDatosProducto.Item(varIndice).proDetalleDatosProductoId = proDatosProducto.proNovedadNumero.Item(varContador).proTipoLinea Then
                                            proDatosProducto.proDetalleDatosProducto.Item(varIndice).FunGDisminuirContador
                                            Exit For
                                        End If
                                    Next
                                End If
                            Else
                                'Incrementar contador de tipo de línea actual
                                varIndice = -1
                                varIndice = proDatosProducto.proDetalleDatosProducto.IndexOf(proDatosProducto.proNovedadNumero.Item(varContador).proTipoLineaAnterior)
                                If varIndice >= 0 Then
                                    proDatosProducto.proDetalleDatosProducto.Item(varIndice).FunGAumentarContador
                                    proDatosProducto.proDetalleDatosProducto.Item(varIndice).proEliminar = False
                                    proDatosProducto.proDetalleDatosProducto.Item(varIndice).proBackUp = False
                                    proDatosProducto.proDetalleDatosProducto.Item(varIndice).proModificar = False
                                End If
                                'Eliminar novedad de tipo de línea en edición, si existe
                                For varIndice = 1 To proDatosProducto.proNovedadDetalleDatosProducto.Count
                                    If proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).proNovedadDetalleDatosProductoId = proDatosProducto.proNovedadNumero.Item(varContador).proTipoLinea Then
                                        proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).MetEliminar
                                        varTiposLinea = False 'Consultar nuevamente tipos de línea
                                        Exit For
                                    End If
                                Next
                            End If
                            'Liberar el número para aquellos que son nuevos
                            If proDatosProducto.proNovedadNumero.Item(varContador).proTipoLineaAnterior = 0 Then
                                varFNumero.proUpdateDate = Format(Now, "mm/dd/yyyy hh:mm:ss")
                                varFNumero.proRegionCode = Me.proDatosProducto.proNovedadNumero.Item(varContador).proRegionCode
                                varFNumero.proNumero = Me.proDatosProducto.proNovedadNumero.Item(varContador).proNumero
                                varNuevoEstado = "L"
                                If varDescripcionFAXTOMAIL <> "" Then
                                    If InStr(1, Me.proDatosProducto.proNovedadNumero.Item(varContador).proClasificacionDescripcion, varDescripcionFAXTOMAIL, vbTextCompare) > 0 Then
                                        varNuevoEstado = "F"
                                    End If
                                End If
                                ' 20070815 - CC:
                                If varDescripcionTELEFONOVIRTUAL <> "" Then
                                    If InStr(1, Me.proDatosProducto.proNovedadNumero.Item(varContador).proClasificacionDescripcion, varDescripcionTELEFONOVIRTUAL, vbTextCompare) > 0 Then
                                        varNuevoEstado = "V"
                                    End If
                                End If
                                varFNumero.proEstadoNumero = varNuevoEstado
                                If varFNumero.FunGModificar Then
                                    
                                Else
                                    MsgBox "Error al actualizar el estado del número.", vbCritical, App.Title
                                End If
                            End If
                            Me.proDatosProducto.proNovedadNumero.Remove (varContador)
                            Me.proDatosProducto.proNovedadNumero.proSeleccionados = Me.proDatosProducto.proNovedadNumero.proSeleccionados - 1
                        Else
                            MsgBox "Error al eliminar el número [" + Me.proDatosProducto.proNovedadNumero.Item(varContador).proNumero + "].", vbCritical, App.Title
                        End If
                    Else
                        varContador = varContador + 1
                    End If
                Else
                    varContador = varContador + 1
                End If
            Wend
            
            Call SubFPintarGridEdicionNumeroPublico(Index)
            
            Me.cmdDeshacerModificación(Index).Enabled = False
        End If
    End If
        cmdModificarInsertados(2).Enabled = varModificarNumeracionPublica And (grdEdicionNumeroPublico.Rows - grdEdicionNumeroPublico.FixedRows > 0 Or grdNumeroPublico.Rows - grdNumeroPublico.FixedRows > 0)
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdEliminar_Click(Index As Integer)
    Dim varContador As Integer, varIndiceDetalleDatos As Integer, varIndiceTipoLinea As Integer, varIndice As Integer
    Dim varContadorAux As Integer
    Dim varEncontro As Boolean
    Dim varContadorSeleccionados As Integer
    Dim varNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
    Dim varNovedadNumeros As claNovedadNumero
    Dim varNovedadNumeracionCorporativa As claNovedadNumeracionCorporativa
    Dim varNuevoTipoLinea As Long
    Dim varBackup As Boolean
    On Error GoTo ErrManager

    'Tipos de líneas
    If Index = 0 Then
        If Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los detalles que desea eliminar.", vbInformation, App.Title
            Exit Sub
        End If
        
        If MsgBox("Desea eliminar los [" & Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados & "] detalles selecionados?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        
            'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
            If Not Me.proDatosProducto.MetGuardar Then
                MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Inserta o actualiza la información de los incidentes
            If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
                MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
            End If
            
            varContadorSeleccionados = 0
            For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
                If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "1" Then
                    varContadorSeleccionados = varContadorSeleccionados + 1
                    
                    'Validar que ya no se encuentre dentro de los seleccionados para eliminar
                    varEncontro = False
                    For varContadorAux = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                        If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId = _
                           Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorAux).proDetalleDatosProductoId _
                           And (Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorAux).proTipoNovedadId = 3 Or Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorAux).proTipoNovedadId = 2) Then
                            varEncontro = True
                            Exit For
                        End If
                    Next varContadorAux
                    
                    If varEncontro = True Then
                        MsgBox "El [" & varContadorSeleccionados & "] registro seleccionado, ya se encuentra marcado para eliminación o modificación."
                    Else
                        Set varNovedadDetalleDatosProducto = New claNovedadDetalleDatosProducto
                        Set varNovedadDetalleDatosProducto.proConexion = Me.proConexion
                        Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proEliminar = True
                        Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proContador = 0
                        Call SubFAsignarRegistroAModificar(Me.proDatosProducto.proDetalleDatosProducto.Item(varContador), varNovedadDetalleDatosProducto, 3)
                        
                        If varNovedadDetalleDatosProducto.MetGuardar Then
                            If Not Me.proDatosProducto.MetAgregarNovedadDetalle(varNovedadDetalleDatosProducto) Then
                                MsgBox "Error al agregar el detalle.", vbCritical, App.Title
                                Exit Sub
                            End If
                        Else
                            MsgBox "Error al guardar el detalle.", vbCritical, App.Title
                            Exit Sub
                        End If
                        
                        'Eliminar los números asignados con el tipo de linea eliminado
                        varNuevoTipoLinea = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proDatosProducto.proNovedadDetalleDatosProducto.Count).proNovedadDetalleDatosProductoId
                        For varIndice = 1 To proDatosProducto.proDatosProductoNumero.Count
                            If proDatosProducto.proDatosProductoNumero.Item(varIndice).proTipoLinea = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId Then
                                
                                'Consultar números en proceso de instalación o modificación
                                If Me.proDatosProducto.proNovedadNumero Is Nothing Then
                                    Set Me.proDatosProducto.proNovedadNumero = New colNovedadNumero
                                    Set Me.proDatosProducto.proNovedadNumero.proConexion = Me.proConexion
                                    
                                    If Me.proDatosProducto.MetConsultarNovedadNumeros Then
                                        Call SubFPintarGridEdicionNumeroPublico(Index)
                                    Else
                                        MsgBox "Error al consultar los números  que se encuentran en proceso de instalación o modificación.", vbCritical, App.Title
                                        Exit Sub
                                    End If
                                End If
                                
                                'Validar que ya no se encuentre dentro de los seleccionados para eliminar
                                varEncontro = False
                                For varContadorAux = 1 To Me.proDatosProducto.proNovedadNumero.Count
                                    If Me.proDatosProducto.proDatosProductoNumero.Item(varIndice).proRegionCode = _
                                       Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proRegionCode _
                                       And Me.proDatosProducto.proDatosProductoNumero.Item(varIndice).proNumero = _
                                       Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proNumero _
                                        Then
                                        If Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proTipoNovedadId = 2 Then
                                          Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proTipoNovedadId = 3
                                          Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proTipoLineaAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId
                                          Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proTipoLinea = varNuevoTipoLinea
                                          Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).FunGEliminarServiciosxReserva
                                           Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).FunGDisminuirContador
                                        End If
                                        varEncontro = True
                                        Exit For
                                    End If
                                Next varContadorAux
                                If Not varEncontro Then
                                    Set varNovedadNumeros = Nothing
                                    Set varNovedadNumeros = New claNovedadNumero
                                    Set varNovedadNumeros.proConexion = Me.proConexion
                                    varNovedadNumeros.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                                    varNovedadNumeros.proFechaLiberacion = ""
                                    varNovedadNumeros.proFechaReserva = ""
                                    varNovedadNumeros.proIncidentId = Me.proDatosProducto.proIncidentId
                                    varNovedadNumeros.proNovedadNumeroId = "0"
                                    varNovedadNumeros.proNumero = Me.proDatosProducto.proDatosProductoNumero.Item(varIndice).proNumero
                                    varNovedadNumeros.proRegionCode = Me.proDatosProducto.proDatosProductoNumero.Item(varIndice).proRegionCode
                                    varNovedadNumeros.proRegionName = Me.proDatosProducto.proDatosProductoNumero.Item(varIndice).proRegionName
                                    varNovedadNumeros.proTipoNovedadId = "3"
                                    varNovedadNumeros.proPublicar = Me.proDatosProducto.proDatosProductoNumero.Item(varIndice).proPublicar
                                    varNovedadNumeros.proTipoLineaAnterior = Me.proDatosProducto.proDatosProductoNumero.Item(varIndice).proTipoLinea
                                    varNovedadNumeros.proTipoLinea = varNuevoTipoLinea
                                    If varNovedadNumeros.FunGInsertar Then
                                        If Not Me.proDatosProducto.MetAgregarNovedadNumeroPublico(varNovedadNumeros) Then
                                            MsgBox "Error al agregar el detalle [" & varNovedadNumeros.proNumero & "].", vbCritical, App.Title
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Error al marcar el número [" & varNovedadNumeros.proNumero & "] para eliminación.", vbInformation, App.Title
                                    End If
                                End If
                            End If
                        Next
                        'Validacion de nuevos numeros relacionados al tipo de linea
                        varIndice = 1
                        If Me.proDatosProducto.proNovedadNumero Is Nothing Then
                            Set Me.proDatosProducto.proNovedadNumero = New colNovedadNumero
                            Set Me.proDatosProducto.proNovedadNumero.proConexion = Me.proConexion
                            If Not Me.proDatosProducto.MetConsultarNovedadNumeros Then
                                MsgBox "Error al consultar los números  que se encuentran en proceso de instalación o modificación.", vbCritical, App.Title
                                Exit Sub
                            End If
                        End If
 
                        While varIndice < proDatosProducto.proNovedadNumero.Count
                            If proDatosProducto.proNovedadNumero.Item(varIndice).proTipoLinea = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId And _
                                  proDatosProducto.proNovedadNumero.Item(varIndice).proTipoNovedadId = 1 Then
                                Dim varFNumero As EDCAdminVoz.claNumero
                                Set varFNumero = New EDCAdminVoz.claNumero
                                Set varFNumero.proConexion = Me.proConexion
                                varFNumero.proUpdateBy = Me.proOnyx.UserLogin
                                varFNumero.proRecordStatus = 1
                                varFNumero.proUpdateDate = Format(Now, "mm/dd/yyyy hh:mm:ss")
                                varFNumero.proRegionCode = Me.proDatosProducto.proNovedadNumero.Item(varIndice).proRegionCode
                                varFNumero.proNumero = Me.proDatosProducto.proNovedadNumero.Item(varIndice).proNumero
                                Dim varNuevoEstado, varDescripcionFAXTOMAIL, varDescripcionTELEFONOVIRTUAL As String
                                varNuevoEstado = "L"
                                If varDescripcionFAXTOMAIL <> "" Then
                                    If InStr(1, Me.proDatosProducto.proNovedadNumero.Item(varIndice).proClasificacionDescripcion, varDescripcionFAXTOMAIL, vbTextCompare) > 0 Then
                                        varNuevoEstado = "F"
                                    End If
                                End If
                                ' 20070815 - CC:
                                If varDescripcionTELEFONOVIRTUAL <> "" Then
                                    If InStr(1, Me.proDatosProducto.proNovedadNumero.Item(varIndice).proClasificacionDescripcion, varDescripcionTELEFONOVIRTUAL, vbTextCompare) > 0 Then
                                        varNuevoEstado = "V"
                                    End If
                                End If
                                varFNumero.proEstadoNumero = varNuevoEstado
                                If varFNumero.FunGModificar Then
                                    Me.proDatosProducto.proNovedadNumero.Item(varIndice).FunGEliminar
                                    Me.proDatosProducto.proNovedadNumero.Remove (varIndice)
                                    Me.proDatosProducto.proNovedadNumero.proSeleccionados = Me.proDatosProducto.proNovedadNumero.proSeleccionados - 1
                                Else
                                    MsgBox "Error al actualizar el estado del número.", vbCritical, App.Title
                                End If
                                varIndice = varIndice - 1
                            End If
                            varIndice = varIndice + 1
                        Wend
                    End If
                    Call SubFPintarGridEdicionNumeroPublico(Index)
                End If
            Next varContador
            varNumerosPublicos = False 'Consultar nuevamente la grilla de números
        End If
        
        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "0"
        Next varContador
    
        Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0
        
        Call SubFPintarGridTiposLinea(Index)
        Call SubFPintarGridTiposLineaModificacion(Index)
        
        Me.cmdModificar(Index).Enabled = False
        Me.cmdModificarColumna(Index).Enabled = False
        Me.cmdEliminar(Index).Enabled = False
        Me.cmdClonar(Index).Enabled = False
    End If

    'Numeración Corporativa
    If Index = 1 Then
        If Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los registros a eliminar.", vbInformation, App.Title
            Exit Sub
        End If
        
        If MsgBox("Desea eliminar los [" & Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados & "] detalles selecionados?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
            If Not Me.proDatosProducto.MetGuardar Then
                MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Inserta o actualiza la información de los incidentes
            If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
                MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
            End If
            
            varContadorSeleccionados = 0
            For varContador = 1 To Me.proDatosProducto.proNumeracionCorporativa.Count
                If Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proSeleccion = "1" Then
                
                    varContadorSeleccionados = varContadorSeleccionados + 1
                    
                    'Validar que ya no se encuentre dentro de los seleccionados para eliminar
                    varEncontro = False
                    For varContadorAux = 1 To Me.proDatosProducto.proNovedadNumeracionCorporativa.Count
                        If Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proMarcacion = _
                           Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContadorAux).proMarcacion _
                           And Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContadorAux).proTipoNovedadId = 3 Then
                           
                            varEncontro = True
                            Exit For
                        End If
                    Next varContadorAux
                    
                    If varEncontro = True Then
                        MsgBox "El [" & varContadorSeleccionados & "] registro seleccionado, ya se encuentra marcado para eliminación."
                    Else
                        Set varNovedadNumeracionCorporativa = Nothing
                        Set varNovedadNumeracionCorporativa = New claNovedadNumeracionCorporativa
                        Set varNovedadNumeracionCorporativa.proConexion = Me.proConexion
                        
                        varNovedadNumeracionCorporativa.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                        varNovedadNumeracionCorporativa.proIncidentId = Me.proDatosProducto.proIncidentId
                        varNovedadNumeracionCorporativa.proMarcacion = Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proMarcacion
                        varNovedadNumeracionCorporativa.proTipoNovedadId = "3"
                        
                        If varNovedadNumeracionCorporativa.FunGInsertar Then
                            If Not Me.proDatosProducto.MetAgregarNovedadNumeracionCorporativa(varNovedadNumeracionCorporativa) Then
                                MsgBox "Error al agregar el detalle [" & varNovedadNumeracionCorporativa.proMarcacion & "].", vbCritical, App.Title
                                Exit Sub
                            End If
                        Else
                            MsgBox "Error al marcar el número [" & varNovedadNumeracionCorporativa.proMarcacion & "] para eliminación.", vbInformation, App.Title
                        End If
                    End If
                End If
            Next varContador
        End If
        
        For varContador = 1 To Me.proDatosProducto.proNumeracionCorporativa.Count
            Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proSeleccion = "0"
        Next varContador
    
        Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados = 0
        
        Call SubFPintarGridNumeracionCorporativa(Index)
        Call SubFPintarGridEdicionNumeracionCorporativa(Index)
        
        Me.cmdEliminar(Index).Enabled = False
        
    End If
    
    'Numeración pública
    If Index = 2 Then
        If Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los registros a eliminar.", vbInformation, App.Title
            Exit Sub
        End If
        If MsgBox("Desea eliminar los [" & Me.proDatosProducto.proDatosProductoNumero.proSeleccionados & "] detalles selecionados?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
            If Not Me.proDatosProducto.MetGuardar Then
                MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
                Exit Sub
            End If
            'Inserta o actualiza la información de los incidentes
            If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
                MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
            End If
            varContadorSeleccionados = 0
            For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
                If Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proSeleccion = "1" Then
                    varContadorSeleccionados = varContadorSeleccionados + 1
                    'Validar que ya no se encuentre dentro de los seleccionados para eliminar
                    varEncontro = False
                    For varContadorAux = 1 To Me.proDatosProducto.proNovedadNumero.Count
                        If Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionCode = _
                           Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proRegionCode _
                           And Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proNumero = _
                           Me.proDatosProducto.proNovedadNumero.Item(varContadorAux).proNumero _
                            Then
                            varEncontro = True
                            Exit For
                        End If
                    Next varContadorAux
                    If varEncontro = True Then
                        MsgBox "El [" & varContadorSeleccionados & "] registro seleccionado, ya se encuentra marcado para eliminación o modificación."
                    Else
                        Set varNovedadNumeros = Nothing
                        Set varNovedadNumeros = New claNovedadNumero
                        Set varNovedadNumeros.proConexion = Me.proConexion
                        varNovedadNumeros.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                        varNovedadNumeros.proFechaLiberacion = ""
                        varNovedadNumeros.proFechaReserva = ""
                        varNovedadNumeros.proIncidentId = Me.proDatosProducto.proIncidentId
                        varNovedadNumeros.proNovedadNumeroId = "0"
                        varNovedadNumeros.proNumero = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proNumero
                        varNovedadNumeros.proRegionCode = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionCode
                        varNovedadNumeros.proRegionName = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionName
                        varNovedadNumeros.proTipoNovedadId = 3
                        varNovedadNumeros.proPublicar = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proPublicar
                        varNovedadNumeros.proTipoLineaAnterior = Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea
                        'Revisar si se debe eliminar el tipo de línea del número eliminado (si no tiene más números asignados)
                        Dim varexiste As Boolean
                        varexiste = False
                        If Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea <> 0 Then 'Si tiene un tipo de línea asignado
                            varIndiceDetalleDatos = proDatosProducto.proDetalleDatosProducto.IndexOf(Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea)
                            If (proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proContador <= 1) Then
                                'Se debe revisar si el tipo de linea no se encuentra en edición
                                Dim varindicenovedad As Integer
                                varindicenovedad = 1
                                While varindicenovedad <= Me.proDatosProducto.proNovedadDetalleDatosProducto.Count And varexiste = False
                                    If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varindicenovedad).proDetalleDatosProductoId = proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proDetalleDatosProductoId Then
                                        varexiste = True
                                    End If
                                    varindicenovedad = varindicenovedad + 1
                                Wend
                                If varexiste = True Then
                                    MsgBox "No es posible eliminar el número ''" & varNovedadNumeros.proNumero & "'' ya que el tipo de línea [" & Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea & "] se encuentra en edición.", vbOKOnly, App.Title
                                Else
                                    'Si se debe eliminar
                                    varBackup = False
                                    varIndiceTipoLinea = varValoresCampoProductoTipoLinea.BuscarIndiceProValorId(proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proUser1)
                                    If varValoresCampoProductoTipoLinea.Item(varIndiceTipoLinea).proUsual <> 1 Then
                                        varBackup = (MsgBox("El tipo de línea ''" & varValoresCampoProductoTipoLinea.Item(varIndiceTipoLinea).proValorDesc & "'' (" & proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proDetalleDatosProductoId & ") no quedará relacionado a ningún número. ¿Desea modificar este tipo de línea como Backup?", vbQuestion + vbYesNo, App.Title) = vbYes)
                                    End If
                                    If varBackup Then
                                        'Crear un incidente de novedad pasando a backup el tipo de linea
                                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proCampo = "vchuser15"
                                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proCodigos = Me.proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proDetalleDatosProductoId & ","
                                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proIncidentId = Me.proDatosProducto.proIncidentId
                                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proProductNumber = Me.proDatosProducto.proProductNumber
                                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proTabla = "1"
                                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = varSi
                                        Me.proDatosProducto.proNovedadDetalleDatosProducto.MetActualizarColumna
                                        Me.proDatosProducto.MetConsultarNovedadDetalleDatosProducto
                                        varNuevoTipoLinea = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proDatosProducto.proNovedadDetalleDatosProducto.Count).proNovedadDetalleDatosProductoId
                                        Me.proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proBackUp = True
                                    Else
                                        'Crear un incidente marcándolo como a eliminar
                                        MsgBox "El tipo de línea ''" & varValoresCampoProductoTipoLinea.Item(varIndiceTipoLinea).proValorDesc & "'' (" & proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proDetalleDatosProductoId & ") no queda relacionado a ningún número, por lo cual será eliminado.", vbInformation, App.Title
                                        Set varNovedadDetalleDatosProducto = New claNovedadDetalleDatosProducto
                                        Set varNovedadDetalleDatosProducto.proConexion = Me.proConexion
                                        SubFAsignarRegistroAModificar proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos), varNovedadDetalleDatosProducto, 3
                                        varNovedadDetalleDatosProducto.MetGuardar
                                        varNuevoTipoLinea = varNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId
                                        Me.proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).proEliminar = True
                                    End If
                                End If
                            End If
                        End If
                        If varexiste = False Then
                            varTiposLinea = False 'Consultar nuevamente pestaña de tipos de línea
                            If varIndiceDetalleDatos <> 0 Then
                                proDatosProducto.proDetalleDatosProducto.Item(varIndiceDetalleDatos).FunGDisminuirContador
                            End If
                             varNovedadNumeros.proTipoLinea = varNuevoTipoLinea
                            If varNovedadNumeros.FunGInsertar Then
                                If Not Me.proDatosProducto.MetAgregarNovedadNumeroPublico(varNovedadNumeros) Then
                                    MsgBox "Error al agregar el detalle [" & varNovedadNumeros.proNumero & "].", vbCritical, App.Title
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Error al marcar el número [" & varNovedadNumeros.proNumero & "] para eliminación.", vbInformation, App.Title
                            End If
                        End If
                    End If
                End If
            Next varContador
            
        End If
        For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
            Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proSeleccion = "0"
        Next varContador
    
        Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = 0
        
        Call SubFPintarGridNumeroPublico(Index)
        Call SubFPintarGridEdicionNumeroPublico(Index)
        Me.CmdCambiarTipoLinea.Enabled = False '-->3.7.4
        Me.cmdEliminar(Index).Enabled = False
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdEliminarCliente_Click()
On Error GoTo ErrorManager
        
        'Verifica la existencia de datos
        If Val(Me.proDatosProducto.proClienteNacionalId) = 0 Then Exit Sub
        
        If MsgBox("Está seguro de eliminar el vínculo con " & Me.proDatosProducto.proNombreClienteNacional & " para Telefonía Nacional?", vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
        
        Me.proDatosProducto.proClienteNacionalId = ""
        Me.proDatosProducto.proNombreClienteNacional = ""
        If Me.proDatosProducto.MetGuardar Then
            MsgBox "Se eliminó el vínculo existente con Telefonía Nacional", vbInformation, App.Title
            Me.lblIDCliente = Me.proDatosProducto.proClienteNacionalId
            Me.lblCliente = Me.proDatosProducto.proNombreClienteNacional
        End If
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub cmdELiminarClienteLocal_Click()
On Error GoTo ErrorManager
        
        'Verifica la existencia de datos
        If Val(Me.proDatosProducto.proClienteNacionalId) = 0 Then Exit Sub
        
        If MsgBox("Está seguro de eliminar el vínculo con " & Me.proDatosProducto.proNombreClienteNacional & " para Telefonía Nacional?", vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
        
        Me.proDatosProducto.proClienteNacionalId = ""
        Me.proDatosProducto.proNombreClienteNacional = ""
        If Me.proDatosProducto.MetGuardar Then
            MsgBox "Se eliminó el vínculo existente con Telefonía Nacional", vbInformation, App.Title
            Me.lblIDClienteLocal = Me.proDatosProducto.proClienteNacionalId
            Me.lblClienteLocal = Me.proDatosProducto.proNombreClienteNacional
        End If
        Exit Sub
        
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
    Dim varClienteTelefonia As EDCVoz.claClienteTelefonia
    On Error GoTo ErrManager
    
    'Validar que colocaron el Grupo Centrex
    If Trim(Me.txtGrupoCentrex.Text) = "" Then
        MsgBox "Debe colocar el Grupo Centrex.", vbInformation, App.Title
        Exit Sub
    End If
    
    'Validar que colocaron en Call Source
    If Trim(Me.txtCallSource.Text) = "" Then
        MsgBox "Debe colocar el Call Source.", vbInformation, App.Title
        Exit Sub
    End If
    
    'Instanciar una copia de la clase para evitar la visualización
    'erronea de la información
    
    Set varClienteTelefonia = New EDCVoz.claClienteTelefonia
    Set varClienteTelefonia.proConexion = Me.proConexion
    
    varClienteTelefonia.proCompanyId = Me.proOnyx.ContactID
    varClienteTelefonia.proGrupoCentrex = Trim(Me.txtGrupoCentrex.Text)
    varClienteTelefonia.proCallSource = Trim(Me.txtCallSource.Text)
    
    'Validar que el Grupo Centrex sea único para este cliente
    If varClienteTelefonia.MetValidarExistenciaGrupoCentrex Then
        MsgBox "El Grupo Centrex seleccionado ya fue asignado a otro cliente.", vbInformation, App.Title
        Exit Sub
    End If
    
    'Validar que el Call Source sea único para este cliente
    If varClienteTelefonia.MetValidarExistenciaCallSource Then
        MsgBox "El Call Source seleccionado ya fue asignado a otro cliente.", vbInformation, App.Title
        Exit Sub
    End If
    
    'Si pasó todas la validaciones debe almacenar la información
    If varClienteTelefonia.MetGuardar Then
    
        'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
        If Not Me.proDatosProducto.MetGuardar Then
            MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
            Exit Sub
        End If
        
        'Inserta o actualiza la información de los incidentes
        If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
            MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
            Exit Sub
        End If
    
        'Asignar los valores a la clase definitiva
        Me.proClienteTelefonia.proCompanyId = varClienteTelefonia.proCompanyId
        Me.proClienteTelefonia.proGrupoCentrex = varClienteTelefonia.proGrupoCentrex
        Me.proClienteTelefonia.proCallSource = varClienteTelefonia.proCallSource
        
        'Liberar la memoria utilizada en la copia
        Set varClienteTelefonia = Nothing
        MsgBox "La información se almacenó exitosamente.", vbInformation, App.Title
    Else
        MsgBox "Error al almacenar la información del Grupo Centrex y del Call Source.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdGuardarEnvio_Click()
    On Error GoTo ErrManager
    
    Me.proDatosProducto.proEnvioPublicoLocal = Me.chkEnvioPublicoLocal.Value
    Me.proDatosProducto.proEnvioPublicoLD = Me.chkEnvioPublicoLD.Value
    Me.proDatosProducto.proEnvioCorpLocal = Me.chkEnvioCorpLocal.Value
    Me.proDatosProducto.proEnvioCorpLD = Me.chkEnvioCorpLD.Value
    
    If Val(Trim(Me.proDatosProducto.proDatosProductoId)) <> 0 Then
        If Me.proDatosProducto.MetActualizar Then
            MsgBox "La información se modificó exitosamente.", vbInformation, App.Title
        Else
            MsgBox "Error al actualizar la información.", vbCritical, App.Title
        End If
    Else
        MsgBox "La información de los envíos se almacenará cuando se almacene el primer detalle de información." _
        & Chr(13) & "Si no se crea ningún detalle esta información no será almacenada.", vbExclamation, App.Title
        
    End If
    Me.chkEnvioCorpLD.Enabled = False
    Me.chkEnvioCorpLocal.Enabled = False
    Me.chkEnvioPublicoLD.Enabled = False
    Me.chkEnvioPublicoLocal.Enabled = False
    Me.cmdCancelarEnvio.Visible = False
    Me.cmdGuardarEnvio.Visible = False
    Me.cmdModificarEnvio.Visible = True
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdInsertar_Click(Index As Integer)
    Dim varContador As Integer, varIndice As Integer, varNumerosPorAsignar As Integer, varIndiceTipoLinea As Long
    Dim varNovedadNumero As claNovedadNumero
    On Error GoTo ErrManager
    
    ' Inicio 2.0.000 Se valida si el tipo de incidente puede alterar la modificacion
    Dim Conteo As String

    'Inicio Julio Salinas
    'Conteo = proDatosProducto.MetConsultaIncidenteTelefonia(varOperacionOnyx.proTipoIncidente, varOperacionOnyx.proCategoriaIncidente)
    
    'Si el tipo de incidente no posee permiso informa al usuario y termina el proceso.
    'If Conteo = "" Then
        
    '    MsgBox ("El tipo de incidente o la categoria del mismo no puede realizar modificaciones sobre la numeracion, verifique que el incidente sea el correcto")
    '    Exit Sub
    
    'End If
    ' Fin 2.0.000
    'FIn Julio Salinas

    'Si el tab es el de Tipos de líneas
    If Index = 0 Then
        Set frmEdicionDetalleDatos.proConexion = Me.proConexion
        Set frmEdicionDetalleDatos.proDatosProducto = Me.proDatosProducto
        Set frmEdicionDetalleDatos.proOnyx = Me.proOnyx
        Set frmEdicionDetalleDatos.proParametroProducto = Me.proDatosProducto.proParametrosProducto
        Set frmEdicionDetalleDatos.proNovedadDetalleDatosProducto = New claNovedadDetalleDatosProducto
        Set frmEdicionDetalleDatos.proNovedadDetalleDatosProducto.proConexion = Me.proConexion
        frmEdicionDetalleDatos.proNovedadDetalleDatosProducto.proIncidentId = varOperacionOnyx.proIncidente
        frmEdicionDetalleDatos.proNovedadDetalleDatosProducto.proTipoNovedadId = 1
        frmEdicionDetalleDatos.proNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
        frmEdicionDetalleDatos.proInsUpd = "I"
        '* 1.0.100 Inicio Se toma el id del cliente para enviarse al siguiente formulario
         frmEdicionDetalleDatos.proiClienteId = Me.lblIDCliente.Caption
        '* 1.0.100 Fin
        frmEdicionDetalleDatos.Show (vbModal)
        Screen.MousePointer = vbHourglass
        Call SubFPintarGridTiposLineaModificacion(Index)
    End If

    'Si el tab es el de Numeración Corporativa
    If Index = 1 Then
        Set frmEdicionNumeracionCorporativa.proConexion = Me.proConexion
        Set frmEdicionNumeracionCorporativa.proDatosProducto = Me.proDatosProducto
        Set frmEdicionNumeracionCorporativa.proOnyx = Me.proOnyx

        frmEdicionNumeracionCorporativa.Show (vbModal)
        Screen.MousePointer = vbHourglass
        Call SubFPintarGridEdicionNumeracionCorporativa(Index)

    End If

    'Si el tab es el de Numeración Pública
    If Index = 2 Then
        'Cargar tipos de linea en edición
        Dim varTipoLineaEdicion As New colTipoLineaEdicion
        For varIndice = 1 To proDatosProducto.proNovedadDetalleDatosProducto.Count
            If proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).proTipoNovedadId <> "3" Then
                varTipoLineaEdicion.Add proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).proNovedadDetalleDatosProductoId, _
                    proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).proUser1, _
                    proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).proUser15, _
                    proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndice).proContadorNumeros, _
                    True
            End If
        Next
        For varIndice = 1 To proDatosProducto.proDetalleDatosProducto.Count
            If proDatosProducto.proDetalleDatosProducto.Item(varIndice).proEliminar = False And proDatosProducto.proDetalleDatosProducto.Item(varIndice).proBackUp = False And proDatosProducto.proDetalleDatosProducto.Item(varIndice).proModificar = False _
            And proDatosProducto.proDetalleDatosProducto.Item(varIndice).proStatusId = "A" Then
                varTipoLineaEdicion.Add proDatosProducto.proDetalleDatosProducto.Item(varIndice).proDetalleDatosProductoId, _
                    proDatosProducto.proDetalleDatosProducto.Item(varIndice).proUser1, _
                    proDatosProducto.proDetalleDatosProducto.Item(varIndice).proUser15, _
                    proDatosProducto.proDetalleDatosProducto.Item(varIndice).proContador, _
                    False
            End If
        Next
                
        'Consultar valor parametrizado para NO
        Dim varClaParametro As New claParametro
        Set varClaParametro.proConexion = Me.proConexion
        varClaParametro.proAcronimo = "ValorNo"
        varClaParametro.FunGConsultar

        Set varConsultaNumeros = New EDCAdminVoz.claConsultaNumero
        Set varConsultaNumeros.proConexion = Me.proConexion
        Set varConsultaNumeros.proNumeros = varNumeros
        Set varConsultaNumeros.proValoresCampoProducto = varValoresCampoProductoTipoLinea
        Set varConsultaNumeros.proTipoLineaEdicion = varTipoLineaEdicion
        varConsultaNumeros.proNo = IIf(IsNull(varClaParametro.proValor), "", varClaParametro.proValor)
        varConsultaNumeros.proCodCiudad = Me.proDatosProducto.proCodigoRegion

        varConsultaNumeros.MetMostrarVentanaConsulta
        
        Set varNumeros = varConsultaNumeros.proNumeros
        
        If varNumeros Is Nothing Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        If varNumeros.proSeleccionados > 0 Then
            If MsgBox("Desea asignar " & IIf(varNumeros.proSeleccionados = 1, "el número seleccionado", "los [" + CStr(varNumeros.proSeleccionados) + "] números seleccionados") & " al cliente?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then

                'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
                If Not Me.proDatosProducto.MetGuardar Then
                    MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
                    Exit Sub
                End If

                'Inserta o actualiza la información de los incidentes
                If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
                    MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
                End If

                Set varNovedadNumero = New claNovedadNumero
                Set varNovedadNumero.proConexion = Me.proConexion
                varNovedadNumero.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                varNovedadNumero.proFechaLiberacion = ""
                varNovedadNumero.proFechaReserva = Format(Now, "mm/dd/yyyy hh:mm:ss")
                varNovedadNumero.proIncidentId = Me.proDatosProducto.proIncidentId
                varNovedadNumero.proTipoNovedadId = 1   'Insertar
                varNumerosPorAsignar = varNumeros.proSeleccionados
                varIndiceTipoLinea = 1
                Dim varAsociaNovedad As String
                Dim numbersReserve As String
                numbersReserve = ""
                For varContador = 1 To varNumeros.Count
                    If varNumeros.Item(varContador).proSeleccionado = "S" Then
                        If varNumeros.Item(varContador).proEstadoNumero <> "L" _
                            And varNumeros.Item(varContador).proEstadoNumero <> "F" _
                            And varNumeros.Item(varContador).proEstadoNumero <> "V" Then
                            '01/08/2006 Se adiciona condición <> "F" para que también tome los números que fueron reservados para el servicio de FAX TO MAIL
                            '15/08/2007 Se adiciona condición <> "V" para que también tome los números que fueron reservados para el servicio de TELEFONO VIRTUAL PUBLICO -- CC
                            MsgBox "El número [" + varNumeros.Item(varContador).proNumero + "] se encuentra en estado [" + varNumeros.Item(varContador).proEstadoNumeroDescripcion + "]. No puede ser asignado.", vbInformation, App.Title
                        Else
                            'Buscar índice del tipo de línea en edición al que se asigna el número
                            If varConsultaNumeros.proTipoLineaBasico Then 'Tipo de línea básica:
                                varAsociaNovedad = "S"
                                If varNumerosPorAsignar > 0 Then
                                    Do While varIndiceTipoLinea <= proDatosProducto.proNovedadDetalleDatosProducto.Count
                                        If proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndiceTipoLinea).proUser15 = varConsultaNumeros.proNo Then 'No es backup
                                            If proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndiceTipoLinea).proUser1 = varConsultaNumeros.proCodigoTipoLineaBasica Then 'Tipo de línea seleccionado
                                                If proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndiceTipoLinea).proContadorNumeros = 0 Then
                                                    varNumerosPorAsignar = varNumerosPorAsignar - 1
                                                    Exit Do
                                                End If
                                            End If
                                        End If
                                        varIndiceTipoLinea = varIndiceTipoLinea + 1
                                    Loop
                                End If
                            Else
                                If varConsultaNumeros.proSeleccionInstalado Then
                                    varIndiceTipoLinea = Me.proDatosProducto.proDetalleDatosProducto.IndexOf(varConsultaNumeros.proIndiceTipoLineaEdicion)
                                    varAsociaNovedad = "N"
                                Else
                                    varIndiceTipoLinea = Me.proDatosProducto.proNovedadDetalleDatosProducto.IndexOf(varConsultaNumeros.proIndiceTipoLineaEdicion)
                                    varAsociaNovedad = "S"
                                End If
                                
                            End If
                            If varAsociaNovedad = "S" Then
                                If varIndiceTipoLinea > proDatosProducto.proNovedadDetalleDatosProducto.Count Then varIndiceTipoLinea = -1
                                'Aumentar contador de números para los tipos de línea en edición asignados
                                proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndiceTipoLinea).FunGAumentarContador
                            Else
                                If varIndiceTipoLinea > proDatosProducto.proDetalleDatosProducto.Count Then varIndiceTipoLinea = -1
                                'Aumentar contador de números para los tipos de línea en edición asignados
                                proDatosProducto.proDetalleDatosProducto.Item(varIndiceTipoLinea).FunGAumentarContador
                            End If
                            varNovedadNumero.proNumero = varNumeros.Item(varContador).proNumero
                            varNovedadNumero.proRegionCode = varNumeros.Item(varContador).proRegionCode
                            varNovedadNumero.proRegionName = varNumeros.Item(varContador).proRegionCodeDescripcion
                            varNovedadNumero.proAsociaNovedad = varAsociaNovedad
                            If varIndiceTipoLinea <> -1 Then
                                If varAsociaNovedad = "S" Then
                                    varNovedadNumero.proTipoLinea = proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndiceTipoLinea).proNovedadDetalleDatosProductoId
                                    varNovedadNumero.proPublicar = IIf(proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndiceTipoLinea).proUser14 = varConsultaNumeros.proNo, "N", "S")
                                Else
                                    varNovedadNumero.proTipoLinea = proDatosProducto.proDetalleDatosProducto.Item(varIndiceTipoLinea).proDetalleDatosProductoId
                                    varNovedadNumero.proPublicar = IIf(proDatosProducto.proDetalleDatosProducto.Item(varIndiceTipoLinea).proUser14 = varConsultaNumeros.proNo, "N", "S")
                                End If
                            End If
                            If varNovedadNumero.FunGInsertar Then
                                'varNumeros.Item(varContador).proEstadoNumero = "R"
                                'Se cambia la anterior línea 31/01/2006 ya que se debe primero aprobar ese número para ser asignado
                                varNumeros.Item(varContador).proEstadoNumero = "P"
                                varNumeros.Item(varContador).proUpdateBy = Me.proOnyx.UserLogin
                                varNumeros.Item(varContador).proUpdateDate = Format(Now, "mm/dd/yyyy hh:mm:ss")
                                If Not varNumeros.Item(varContador).FunGModificar Then
                                    MsgBox "Error al actualizar el estado del número.", vbCritical, App.Title
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Error al insertar el número [" + varNumeros.Item(varContador).proNumero + "].", vbCritical, App.Title
                                'varNumeros.proSeleccionados = varNumeros.proSeleccionados - 1
                                proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndiceTipoLinea).proContadorNumeros = proDatosProducto.proNovedadDetalleDatosProducto.Item(varIndiceTipoLinea).proContadorNumeros - 1
                                Screen.MousePointer = vbDefault
                                'Exit Sub
                            End If
                        End If
                        If (numbersReserve <> "") Then
                            numbersReserve = numbersReserve & "," & varNumeros.Item(varContador).proNumero
                        Else
                            numbersReserve = varNumeros.Item(varContador).proNumero
                        End If
                    End If
                Next varContador
                
                Set Me.proDatosProducto.proNovedadNumero = Nothing
                Set Me.proDatosProducto.proNovedadNumero = New colNovedadNumero
                Set Me.proDatosProducto.proNovedadNumero.proConexion = Me.proConexion
                Me.proDatosProducto.proNovedadNumero.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                Me.proDatosProducto.proNovedadNumero.proIncidentId = Me.proDatosProducto.proIncidentId
                If Me.proDatosProducto.proNovedadNumero.MetConsultar(proDatosProducto.proDetalleDatosProducto) Then
                    Call SubFPintarGridEdicionNumeroPublico(Index)
                Else
                    MsgBox "Error al consultar los números públicos.", vbCritical, App.Title
                End If
            End If
        End If
        Set varNumeros = Nothing
    End If
    cmdModificarInsertados(2).Enabled = varModificarNumeracionPublica And (grdEdicionNumeroPublico.Rows - grdEdicionNumeroPublico.FixedRows > 0 Or grdNumeroPublico.Rows - grdNumeroPublico.FixedRows > 0)
    Screen.MousePointer = vbDefault
    
    Dim classPeticionWS As claPeticionNetcracker
    Dim resultadoConsult As Object
    Set classPeticionWS = New claPeticionNetcracker
    Set classPeticionWS.proConexion = Me.proConexion
    
    Set resultadoConsult = classPeticionWS.ParametrosPeticionWs("reserveNumbers", "", "", "TCRM", "Example_PMO-001", "Example_PMO-001", "1", "57", "Pruebas PMO Inspira CLARO", "Client 101800000", "CC", "101800000", "Address Client 101800000", numbersReserve, "", "", "", "", "", "", "", "P")
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdLiberarRecursos_Click()
    On Error GoTo ErrManager
    
    Me.proClienteTelefonia.proCompanyId = Me.proOnyx.ContactID
    Set Me.proClienteTelefonia.proConexion = Me.proConexion
    
    If Me.proClienteTelefonia.MetValidarExistenciaCliente Then
        If Me.proClienteTelefonia.MetEliminar Then
            Me.txtGrupoCentrex.Text = ""
            Me.txtCallSource.Text = ""
            MsgBox "Los recursos se liberaron exitosamente.", vbInformation, App.Title
        Else
            MsgBox "Error al liberar los recursos.", vbCritical, App.Title
            Exit Sub
        End If
    Else
        MsgBox "No existen recursos para ser liberados.", vbInformation, App.Title
    End If
    
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdModificar_Click(Index As Integer)
    Dim varContador As Integer
    Dim varContadorAux As Integer
    Dim varEncontro As Boolean
    Dim varNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
    On Error GoTo ErrManager
    
    If Index = 0 Then
        If Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar el detalle que desea modificar.", vbInformation, App.Title
            Exit Sub
        End If
        
        If Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados > 1 Then
            MsgBox "Para ejecutar esta opción solo debe existir un registro seleccionado.", vbInformation, App.Title
            Exit Sub
        End If
        
        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "1" Then
                If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId <> "A" Or Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proEliminar = True Then
                    MsgBox "Solo puede modificar registros activos.", vbInformation, App.Title
                    Exit Sub
                End If
                Exit For
            End If
        Next varContador
        
        Set frmEdicionDetalleDatos.proConexion = Me.proConexion
        Set frmEdicionDetalleDatos.proDatosProducto = Me.proDatosProducto
        Set frmEdicionDetalleDatos.proOnyx = Me.proOnyx
        Set frmEdicionDetalleDatos.proParametroProducto = Me.proDatosProducto.proParametrosProducto
        
        'Validar si el registro seleccionado ya fue modificado por este incidente
        For varContadorAux = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
           If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId = _
              Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorAux).proDetalleDatosProductoId Then
                varEncontro = True
                Exit For
            End If
        Next varContadorAux
        
        
        'Si lo encontro se edita el que esta en modificacion y si no lo encontro se deben
        'asignar las propiedades del registro seleccionado a las propiedades de modificacion
        If varEncontro Then
            Set frmEdicionDetalleDatos.proNovedadDetalleDatosProducto = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorAux)
        Else
            Set varNovedadDetalleDatosProducto = New claNovedadDetalleDatosProducto
            Set varNovedadDetalleDatosProducto.proConexion = Me.proConexion
            Call SubFAsignarRegistroAModificar(Me.proDatosProducto.proDetalleDatosProducto.Item(varContador), varNovedadDetalleDatosProducto, 2)
            Set frmEdicionDetalleDatos.proNovedadDetalleDatosProducto = varNovedadDetalleDatosProducto
        End If
        
        frmEdicionDetalleDatos.proInsUpd = "U"
       '* 1.0.100 Inicio Se toma el id del cliente para enviarse al siguiente formulario
        frmEdicionDetalleDatos.proiClienteId = Me.lblIDCliente.Caption
        '* 1.0.100 Fin
        frmEdicionDetalleDatos.Show (vbModal)
        
        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0
        
        Call SubFPintarGridTiposLinea(Index)
        Call SubFPintarGridTiposLineaModificacion(Index)
        
        Me.cmdModificar(Index).Enabled = False
        Me.cmdModificarColumna(Index).Enabled = False
        Me.cmdEliminar(Index).Enabled = False
        Me.cmdClonar(Index).Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Public Sub SubFAsignarRegistroAModificar(ByRef parDetalleDatosProducto As claDetalleDatosProducto, ByRef parNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto, ByVal parTipoNovedadId As Integer)
    On Error GoTo ErrManager
    
    parNovedadDetalleDatosProducto.proDatosProductoId = parDetalleDatosProducto.proDatosProductoId
    parNovedadDetalleDatosProducto.proDetalleDatosProductoId = parDetalleDatosProducto.proDetalleDatosProductoId
    parNovedadDetalleDatosProducto.proIncidentId = Me.proDatosProducto.proIncidentId
    parNovedadDetalleDatosProducto.proRecordStatus = parDetalleDatosProducto.proRecordStatus
    parNovedadDetalleDatosProducto.proStatusId = parDetalleDatosProducto.proStatusId
    parNovedadDetalleDatosProducto.proTipoNovedadId = parTipoNovedadId
    parNovedadDetalleDatosProducto.proUser1 = parDetalleDatosProducto.proUser1
    parNovedadDetalleDatosProducto.proUser2 = parDetalleDatosProducto.proUser2
    parNovedadDetalleDatosProducto.proUser3 = parDetalleDatosProducto.proUser3
    parNovedadDetalleDatosProducto.proUser4 = parDetalleDatosProducto.proUser4
    parNovedadDetalleDatosProducto.proUser5 = parDetalleDatosProducto.proUser5
    parNovedadDetalleDatosProducto.proUser6 = parDetalleDatosProducto.proUser6
    parNovedadDetalleDatosProducto.proUser7 = parDetalleDatosProducto.proUser7
    parNovedadDetalleDatosProducto.proUser8 = parDetalleDatosProducto.proUser8
    parNovedadDetalleDatosProducto.proUser9 = parDetalleDatosProducto.proUser9
    parNovedadDetalleDatosProducto.proUser10 = parDetalleDatosProducto.proUser10
    parNovedadDetalleDatosProducto.proUser11 = parDetalleDatosProducto.proUser11
    parNovedadDetalleDatosProducto.proUser12 = parDetalleDatosProducto.proUser12
    parNovedadDetalleDatosProducto.proUser13 = parDetalleDatosProducto.proUser13
    parNovedadDetalleDatosProducto.proUser14 = parDetalleDatosProducto.proUser14
    parNovedadDetalleDatosProducto.proUser15 = parDetalleDatosProducto.proUser15
    parNovedadDetalleDatosProducto.proUser16 = parDetalleDatosProducto.proUser16
    parNovedadDetalleDatosProducto.proUser17 = parDetalleDatosProducto.proUser17
    parNovedadDetalleDatosProducto.proUser18 = parDetalleDatosProducto.proUser18
    parNovedadDetalleDatosProducto.proUser19 = parDetalleDatosProducto.proUser19
    parNovedadDetalleDatosProducto.proUser20 = parDetalleDatosProducto.proUser20
    parNovedadDetalleDatosProducto.proUser21 = parDetalleDatosProducto.proUser21
    parNovedadDetalleDatosProducto.proUser22 = parDetalleDatosProducto.proUser22
    parNovedadDetalleDatosProducto.proUser23 = parDetalleDatosProducto.proUser23
    parNovedadDetalleDatosProducto.proUser24 = parDetalleDatosProducto.proUser24
    parNovedadDetalleDatosProducto.proUser25 = parDetalleDatosProducto.proUser25
    parNovedadDetalleDatosProducto.proUser26 = parDetalleDatosProducto.proUser26
    parNovedadDetalleDatosProducto.proUser27 = parDetalleDatosProducto.proUser27
    parNovedadDetalleDatosProducto.proUser28 = parDetalleDatosProducto.proUser28
    parNovedadDetalleDatosProducto.proUser29 = parDetalleDatosProducto.proUser29
    parNovedadDetalleDatosProducto.proUser30 = parDetalleDatosProducto.proUser30
    parNovedadDetalleDatosProducto.proUser31 = parDetalleDatosProducto.proUser31
    parNovedadDetalleDatosProducto.proUser32 = parDetalleDatosProducto.proUser32
    parNovedadDetalleDatosProducto.proUser33 = parDetalleDatosProducto.proUser33
    parNovedadDetalleDatosProducto.proUser34 = parDetalleDatosProducto.proUser34
    parNovedadDetalleDatosProducto.proUser35 = parDetalleDatosProducto.proUser35
    parNovedadDetalleDatosProducto.proUser36 = parDetalleDatosProducto.proUser36
    parNovedadDetalleDatosProducto.proUser37 = parDetalleDatosProducto.proUser37
    parNovedadDetalleDatosProducto.proUser38 = parDetalleDatosProducto.proUser38
    parNovedadDetalleDatosProducto.proUser39 = parDetalleDatosProducto.proUser39
    parNovedadDetalleDatosProducto.proUser40 = parDetalleDatosProducto.proUser40
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdModificarColumna_Click(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    'Tipos de líneas
    If Index = 0 Then
        If Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los registros a los que desea modificarles la columna.", vbInformation, App.Title
            Exit Sub
        End If
        
        Set frmModificarColumna.proConexion = Me.proConexion
        Set frmModificarColumna.proDatosProducto = Me.proDatosProducto
        Set frmModificarColumna.proOnyx = Me.proOnyx
        frmModificarColumna.proOrigen = "A"
        
        frmModificarColumna.Show (vbModal)
        
        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0
        
        Call SubFPintarGridTiposLinea(Index)
        Call SubFPintarGridTiposLineaModificacion(Index)
        
        Me.cmdModificar(Index).Enabled = False
        Me.cmdModificarColumna(Index).Enabled = False
        Me.cmdEliminar(Index).Enabled = False
        Me.cmdClonar(Index).Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdModificarColumnaInsertados_Click(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    'Tipos de líneas
    If Index = 0 Then
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar los registros a los que desea modificarles la columna.", vbInformation, App.Title
            Exit Sub
        End If
        
        Set frmModificarColumna.proConexion = Me.proConexion
        Set frmModificarColumna.proDatosProducto = Me.proDatosProducto
        Set frmModificarColumna.proOnyx = Me.proOnyx
        frmModificarColumna.proOrigen = "I"
        
        frmModificarColumna.Show (vbModal)
        
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0
        
        Call SubFPintarGridTiposLineaModificacion(Index)
        
        Me.cmdClonarModificados(Index).Enabled = False
        Me.cmdDeshacerModificación(Index).Enabled = False
        Me.cmdModificarInsertados(Index).Enabled = False
        Me.cmdModificarColumnaInsertados(Index).Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
    End Sub

Private Sub cmdModificarEnvio_Click()
    On Error GoTo ErrManager
    
    Me.chkEnvioCorpLD.Enabled = True
    Me.chkEnvioCorpLocal.Enabled = True
    Me.chkEnvioPublicoLD.Enabled = True
    Me.chkEnvioPublicoLocal.Enabled = True
    
    Me.cmdCancelarEnvio.Visible = True
    Me.cmdGuardarEnvio.Visible = True
    Me.cmdGuardarEnvio.Enabled = False
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdModificarInsertados_Click(Index As Integer)
    Dim varContador As Integer
    Dim varContadorAux As Integer
    Dim varEncontro As Boolean
    Dim varNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
    On Error GoTo ErrManager
    
    'Tipos e líneas
    If Index = 0 Then
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0 Then
            MsgBox "Debe seleccionar el detalle que desea modificar.", vbInformation, App.Title
            Exit Sub
        End If
        
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados > 1 Then
            MsgBox "Para ejecutar esta opción solo debe existir un registro seleccionado.", vbInformation, App.Title
            Exit Sub
        End If
        
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "1" Then
                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proTipoNovedadId <> 1 Then
                    MsgBox "Solo puede modificar registros que se están insertando. Los demás registros debe modificarlos en la sección de registros actuales.", vbInformation, App.Title
                    Exit Sub
                End If
                Exit For
            End If
        Next varContador
        
        Set frmEdicionDetalleDatos.proConexion = Me.proConexion
        Set frmEdicionDetalleDatos.proDatosProducto = Me.proDatosProducto
        Set frmEdicionDetalleDatos.proOnyx = Me.proOnyx
        Set frmEdicionDetalleDatos.proParametroProducto = Me.proDatosProducto.proParametrosProducto
        
        
        'Si lo encontro se edita el que esta en modificacion y si no lo encontro se deben
        'asignar las propiedades del registro seleccionado a las propiedades de modificacion
        Set frmEdicionDetalleDatos.proNovedadDetalleDatosProducto = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador)
        
        frmEdicionDetalleDatos.proInsUpd = "U"
      '* 1.0.100 Inicio Se toma el id del cliente para enviarse al siguiente formulario
        frmEdicionDetalleDatos.proiClienteId = Me.lblIDCliente.Caption
        '* 1.0.000 Fin
        frmEdicionDetalleDatos.Show (vbModal)
        
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "0"
        Next varContador
        
        Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0
        
        Call SubFPintarGridTiposLineaModificacion(Index)
        
        Me.cmdClonarModificados(Index).Enabled = False
        Me.cmdDeshacerModificación(Index).Enabled = False
        Me.cmdModificarInsertados(Index).Enabled = False
        Me.cmdModificarColumnaInsertados(Index).Enabled = False
    End If
    
    'Numeración Corporativa
    If Index = 1 Then
    
    End If
    
    'Numeración Pública
    If Index = 2 Then
        'Pasar la conexion
        Set frmEdicionServicios.proConexion = Me.proConexion
        Set frmEdicionServicios.proDatosProducto = Me.proDatosProducto
     
        'Abrir la ventana de edicion
        frmEdicionServicios.Show vbModal
        SubFPintarGridEdicionNumeroPublico (2)
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub CmdModificarLocal_Click()
 On Error GoTo ErrManager
 
    'Validar campos requeridos
    If cboEstratos.ListIndex < 0 Or cboUso.ListIndex < 0 Then
        MsgBox "Debe configurar el estrato y uso de servicio.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Actualizar en la clase valores modificados
    Me.proDatosProducto.proiEstratoid = cboCodigoEstracto.Text: Me.proDatosProducto.proDescripcionEstrato = cboEstratos.Text
    Me.proDatosProducto.proUsoServicioId = cboCodigoUso.Text: Me.proDatosProducto.proDescripcionUso = cboUso.Text
    
    'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
    If Not Me.proDatosProducto.MetGuardar Then
        MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Inserta o actualiza la información de los incidentes
    If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
        MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
        Exit Sub
    End If
    
    If Me.proDatosProducto.proiVentaid <> Trim(Me.TxtIdVenta.Text) Then
        If FunGValidarAsuntoTelefonia(Me.TxtIdVenta.Text, Me.proConexion) = False Then Exit Sub
        Me.proDatosProducto.proiVentaid = Trim(Me.TxtIdVenta.Text)
        Me.proDatosProducto.MetActualizar
    End If
     If varInsertarTiposLinea And Me.proDatosProducto.proiEstratoid <> "" Then
       Me.cmdInsertar(0).Enabled = True
     Else
       Me.cmdInsertar(0).Enabled = False
     End If
    MsgBox "La modificación se ha realizado con exito.", vbOKOnly, App.Title
Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdPublicar_Click()
    Dim varContador As Integer
    On Error GoTo ErrManager
    For varContador = 1 To Me.proDatosProducto.proNovedadNumero.Count
        If Me.proDatosProducto.proNovedadNumero.Item(varContador).proSeleccion = "1" Then
            Me.proDatosProducto.proNovedadNumero.Item(varContador).proPublicar = IIf(Me.proDatosProducto.proNovedadNumero.Item(varContador).proPublicar = "S", "N", "S")
            Me.proDatosProducto.proNovedadNumero.Item(varContador).FunGModificar
        End If
    Next varContador
    Call SubFPintarGridEdicionNumeroPublico(2)
Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdRefrescarPlanesNumeracion_Click()
    On Error GoTo ErrManager
    
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
    
    'Consultar del Plan de Numeración En Curso
    Set varPlanNumeracionEnCurso = New colPlanNumeracionEnCurso
    Set varPlanNumeracionEnCurso.proConexion = Me.proConexion
    varPlanNumeracionEnCurso.proCliente = Me.proOnyx.ContactID
    
    If varPlanNumeracionEnCurso.MetConsultaEnCurso Then
        Call SubFPintarTreePlanNumeracionEnCurso
    Else
        MsgBox "Error al consultar el plan de numeración en curso de instalación.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdSeleccionarTodos_Click(Index As Integer)
    Dim varContador As Integer
    Dim varSeleccionados As Integer
    On Error GoTo ErrManager:
            
    Screen.MousePointer = 11
    
    varSeleccionados = 0
    
    'Tipos de líneas
    Select Case Index
        Case 0
            For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
                If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId = "A" Then
                    Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "1"
                    varSeleccionados = varSeleccionados + 1
                End If
            Next varContador
            
            Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = varSeleccionados
            
            Call SubFPintarGridTiposLinea(Index)
            
            If varSeleccionados > 0 Then
                If varModificarTiposLinea Then
                    Me.cmdModificarColumna(Index).Enabled = True
                    Me.cmdModificar(Index).Enabled = True
                End If
                
                If varInsertarTiposLinea Then
                    Me.cmdClonar(Index).Enabled = True
                End If
                
                If varEliminarTiposLinea Then
                    Me.cmdEliminar(Index).Enabled = True
        
                End If
            End If
        Case 1
            For varContador = 1 To Me.proDatosProducto.proNumeracionCorporativa.Count
                Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proSeleccion = "1"
                varSeleccionados = varSeleccionados + 1
            Next varContador
            
            Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados = varSeleccionados
            
            Call SubFPintarGridNumeracionCorporativa(Index)
            
            If varSeleccionados > 0 Then
                If varModificarNumeracionCorporativa Then
                    Me.cmdModificar(Index).Enabled = True
                End If
                
                If varEliminarNumeracionCorporativa Then
                    Me.cmdEliminar(Index).Enabled = True
                End If
                
            End If
        Case 2
            For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
                Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proSeleccion = "1"
                varSeleccionados = varSeleccionados + 1
            Next varContador
            
            Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = varSeleccionados
            
            Call SubFPintarGridNumeroPublico(Index)
            
            If varSeleccionados > 0 Then
                If varModificarNumeracionPublica Then
                    Me.cmdModificar(Index).Enabled = True
                End If
        
                If varEliminarNumeracionPublica Then
                    Me.cmdEliminar(Index).Enabled = True
                End If
            End If
            Me.CmdCambiarTipoLinea.Enabled = True '-->3.7.4
    End Select
    
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub

Private Sub cmdSeleccionarTodosModificacion_Click(Index As Integer)
    Dim varContador As Integer
    Dim varSeleccionados As Integer
    On Error GoTo ErrManager:
    
    'Tipos de líneas
    Select Case Index
        Case 0
            varSeleccionados = 0
            For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "1"
                varSeleccionados = varSeleccionados + 1
            Next varContador
            
            Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = varSeleccionados
            
            Call SubFPintarGridTiposLineaModificacion(Index)
            
            If varInsertarTiposLinea Then
                Me.cmdClonarModificados(Index).Enabled = True
            End If
            
            If varModificarTiposLinea = True Then
                Me.cmdModificarColumnaInsertados(Index).Enabled = True
                Me.cmdDeshacerModificación(Index).Enabled = True
                Me.cmdModificarInsertados(Index).Enabled = True
            End If
            
        Case 1
            varSeleccionados = 0
            For varContador = 1 To Me.proDatosProducto.proNovedadNumeracionCorporativa.Count
                Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proSeleccion = "1"
                varSeleccionados = varSeleccionados + 1
            Next varContador
            
            Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados = varSeleccionados
            
            Call SubFPintarGridEdicionNumeracionCorporativa(Index)
            
            If varModificarNumeracionCorporativa And Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados <> 0 Then
                Me.cmdDeshacerModificación(Index).Enabled = True
            End If
        Case 2
            varSeleccionados = 0
            For varContador = 1 To Me.proDatosProducto.proNovedadNumero.Count
                Me.proDatosProducto.proNovedadNumero.Item(varContador).proSeleccion = "1"
                varSeleccionados = varSeleccionados + 1
            Next varContador
            
            Me.proDatosProducto.proNovedadNumero.proSeleccionados = varSeleccionados
            
            Call SubFPintarGridEdicionNumeroPublico(Index)
            
            If varModificarNumeracionPublica And Me.proDatosProducto.proNovedadNumero.proSeleccionados <> 0 Then
                Me.cmdDeshacerModificación(Index).Enabled = True
            End If
            
    End Select
   
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub




Private Sub Form_Load()
    Dim varColCliente As EDCTraslados.colCliente
    On Error GoTo ErrManager
        
        Me.TbFondo.Tab = 0
        TbFondo_Click 0
        
        'Inicializar la información del grupo centrex y del call source
        Me.pnlGrupoCentrexCallSource.Visible = False
        
        Me.txtGrupoCentrex.Text = Me.proClienteTelefonia.proGrupoCentrex
        Me.txtCallSource.Text = Me.proClienteTelefonia.proCallSource

        varTiposLinea = False
        varNumerosPrivados = False
        varNumerosPublicos = False
        
        Me.txtCodigo.Text = Me.proDatosProducto.proDatosProductoId
        Me.txtComentarios.Text = Me.proDatosProducto.proComentarios
        
        If Trim(Me.proDatosProducto.proClienteNacionalId) = "" Or Trim(Me.proDatosProducto.proClienteNacionalId) = "0" Then
            Me.proDatosProducto.proClienteNacionalId = Me.proOnyx.ContactID
            Me.proDatosProducto.proNombreClienteNacional = Me.proOnyx.ContactName
        End If

        Set varColCliente = New EDCTraslados.colCliente
        Set varColCliente.proConexion = Me.proConexion

        varColCliente.proClienteId = Me.proDatosProducto.proClienteNacionalId
        If varColCliente.funGConsultaClientexID = True Then
            Me.lblCiudad = varColCliente.Item(1).proCiudad
            Me.lblDireccion = varColCliente.Item(1).proDireccion
            Me.lblsede = varColCliente.Item(1).proSede
        Else
                MsgBox "No fue posible encontrar el cliente " & Me.proDatosProducto.proClienteNacionalId
                Exit Sub
        End If

        If Trim(Me.proDatosProducto.proClienteLocalId) <> "" And Trim(Me.proDatosProducto.proClienteLocalId) <> "0" Then
            Set varColCliente = New EDCTraslados.colCliente
            Set varColCliente.proConexion = Me.proConexion
            varColCliente.proClienteId = Me.proDatosProducto.proClienteLocalId

            If varColCliente.funGConsultaClientexID = True Then
                Me.lblCiudadlocal = varColCliente.Item(1).proCiudad
                Me.lblDireccionLocal = varColCliente.Item(1).proDireccion
                Me.lblSedeLocal = varColCliente.Item(1).proSede
                Me.lblClienteLocal = varColCliente.Item(1).proNombreCliente
                Me.lblIDClienteLocal = varColCliente.Item(1).proClienteId
            Else
                MsgBox "No fue posible encontrar el cliente " & Me.proDatosProducto.proClienteLocalId
                Exit Sub
            End If
        End If

        Me.lblIDCliente = Me.proDatosProducto.proClienteNacionalId
        Me.lblCliente = Me.proDatosProducto.proNombreClienteNacional

        If Me.proDatosProducto.proEnvioPublicoLocal = "True" Or Me.proDatosProducto.proEnvioPublicoLocal = "1" Then
            Me.chkEnvioPublicoLocal.Value = 1
        Else
            Me.chkEnvioPublicoLocal.Value = 0
        End If
        If Me.proDatosProducto.proEnvioPublicoLD = "True" Or Me.proDatosProducto.proEnvioPublicoLD = "1" Then
            Me.chkEnvioPublicoLD.Value = 1
        Else
            Me.chkEnvioPublicoLD.Value = 0
        End If
        
        If Me.proDatosProducto.proEnvioCorpLocal = "True" Or Me.proDatosProducto.proEnvioCorpLocal = "1" Then
            Me.chkEnvioCorpLocal.Value = 1
        Else
            Me.chkEnvioCorpLocal.Value = 0
        End If
        
        If Me.proDatosProducto.proEnvioCorpLD = "True" Or Me.proDatosProducto.proEnvioCorpLD = "1" Then
            Me.chkEnvioCorpLD.Value = 1
        Else
            Me.chkEnvioCorpLD.Value = 0
        End If
        Me.TxtIdVenta.Text = Trim(Me.proDatosProducto.proiVentaid)
        Me.TxtEnlace.Text = Trim(Me.proDatosProducto.proCodigoEnlace)
        Call SubFInicializarBotones
        
        'Consultar ciudad de instalación
        If Me.proDatosProducto.proCiudadId = 0 Then Me.proDatosProducto.MetConsultarCiudadDestino 'Consultar la ciudad, al editar por primera vez
        lblCiudadInstalacion = Me.proDatosProducto.proNombreCiudad

        'Llenar el combo de tipos de estratos
        Call SubFLlenarComboEstratos
        Dim varContadorAux As Integer
        If Len(Me.proDatosProducto.proDescripcionEstrato) > 0 Then
            For varContadorAux = 0 To Me.cboEstratos.ListCount - 1
                If Me.cboCodigoEstracto.List(varContadorAux) = Me.proDatosProducto.proiEstratoid Then
                    Me.cboCodigoEstracto.ListIndex = varContadorAux
                    Me.cboEstratos.ListIndex = varContadorAux
                    Exit For
                End If
            Next
        End If

        'Llenar el combo de usos
        Call SubFLlenarComboUsos
        If Len(Me.proDatosProducto.proDescripcionUso) > 0 Then
            For varContadorAux = 0 To cboUso.ListCount - 1
                If cboCodigoUso.List(varContadorAux) = Me.proDatosProducto.proUsoServicioId Then
                    cboCodigoUso.ListIndex = varContadorAux
                    cboUso.ListIndex = varContadorAux
                    Exit For
                End If
            Next
        End If

        'Deshabilitar los campos estrato y uso para registro de novedades
        cboEstratos.Enabled = True: cboUso.Enabled = True
        
        'Deshabilitar los campos estrato y uso para ventas cerradas
        Dim varProceso As New claProceso
        Set varProceso.proConexion = Me.proConexion
        varProceso.proIncidentId = Me.proDatosProducto.proIncidentId
        If Not varProceso.MetValidarOTCerrada(False) Then
            cboEstratos.Enabled = False: cboUso.Enabled = False
        End If

        'Consultar tipos de línea
        Set varValoresCampoProductoTipoLinea = New EDCAdminVoz.colValoresCampoProducto
        Set varValoresCampoProductoTipoLinea.proConexion = Me.proConexion
        varValoresCampoProductoTipoLinea.proProductNumber = "1810"
        varValoresCampoProductoTipoLinea.proCampo = "vchuser1"
        varValoresCampoProductoTipoLinea.MetConsultarValoresxProducto
        
        'Consultar parámetros
        Dim varClaParametro As New claParametro
        Set varClaParametro.proConexion = Me.proConexion
        varClaParametro.proAcronimo = "ValorSi"
        varClaParametro.FunGConsultar
        varSi = IIf(IsNull(varClaParametro.proValor), "", varClaParametro.proValor)
        
        
        
        Me.CmdCambiarTipoLinea.Enabled = False '-->3.7.4
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarBotones()
    Dim varContador As Integer
    Dim varProceso As claProceso
    On Error GoTo ErrManager
    
    Set varOperacionOnyx = New EDCAdminVoz.colOperacionOnyx
    Set varOperacionOnyx.proConexion = Me.proConexion
    
    Set varProceso = New claProceso
    Set varProceso.proConexion = Me.proConexion
    
    varProceso.proIncidentId = Me.proDatosProducto.proIncidentId
    
    If varProceso.MetConsultaDatosIncidente Then
        If varProceso.proOTId = "0" Or Trim(varProceso.proOTId) = "" Then
            varOperacionOnyx.proIncidente = Me.proDatosProducto.proIncidentId
        Else
            varOperacionOnyx.proIncidente = varProceso.proOTId
        End If
    Else
        MsgBox "Error al validar la OT del incidente.", vbCritical, App.Title
        Exit Sub
    End If
    
    varInsertarTiposLinea = False
    varModificarTiposLinea = False
    varEliminarTiposLinea = False
    varInsertarNumeracionCorporativa = False
    varModificarNumeracionCorporativa = False
    varEliminarNumeracionCorporativa = False
    varInsertarNumeracionPublica = False
    varModificarNumeracionPublica = False
    varEliminarNumeracionPublica = False

    If varOperacionOnyx.MetConsultarxTipoCategoria Then
        For varContador = 1 To varOperacionOnyx.Count
            Select Case varOperacionOnyx.Item(varContador).proTipoSeccionId
                Case "T"
                    Select Case varOperacionOnyx.Item(varContador).proTipoNovedadId
                        Case 1
                            varInsertarTiposLinea = True
                        Case 2
                            varModificarTiposLinea = True
                        Case 3
                            varEliminarTiposLinea = True
                    End Select
                Case "C"
                    Select Case varOperacionOnyx.Item(varContador).proTipoNovedadId
                        Case 1
                            varInsertarNumeracionCorporativa = True
                        Case 2
                            varModificarNumeracionCorporativa = True
                        Case 3
                            varEliminarNumeracionCorporativa = True
                    End Select
                Case "P"
                    Select Case varOperacionOnyx.Item(varContador).proTipoNovedadId
                        Case 1
                            varInsertarNumeracionPublica = True
                        Case 2
                            varModificarNumeracionPublica = True
                        Case 3
                            varEliminarNumeracionPublica = True
                    End Select
                Case "*"
                    Select Case varOperacionOnyx.Item(varContador).proTipoNovedadId
                        Case 1
                            varInsertarTiposLinea = True
                            varInsertarNumeracionCorporativa = True
                            varInsertarNumeracionPublica = True
                        Case 2
                            varModificarTiposLinea = True
                            varModificarNumeracionCorporativa = True
                            varModificarNumeracionPublica = True
                        Case 3
                            varEliminarTiposLinea = True
                            varEliminarNumeracionCorporativa = True
                            varEliminarNumeracionPublica = True
                    End Select
            End Select
        Next varContador
    Else
        MsgBox "Error al consultar los tipo de atencion validas.", vbCritical, App.Title
        Exit Sub
    End If
        
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarGridTiposLinea(Index As Integer)
    Dim varContador As Integer
    Dim varValor As String
    Dim varValorLista As EDCAdminVoz.claValor
    Dim varCantidadRegistros As Integer
    Dim varContadorAux As Integer
    Dim varValorCampo As String
    On Error GoTo ErrManager
    
    'Tipos de líneas
    If Index = 0 Then
    
        If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
            Exit Sub
        End If
        
        varValor = ""
        Me.grdDetalles.Redraw = False
        Me.grdDetalles.Rows = 1
        varCantidadRegistros = 0
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
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1
                            Case "vchUser2"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2
                            Case "vchUser3"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3
                            Case "vchUser4"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4
                            Case "vchUser5"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5
                            Case "vchUser6"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6
                            Case "vchUser7"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7
                            Case "vchUser8"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8
                            Case "vchUser9"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9
                            Case "vchUser10"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10
                            Case "vchUser11"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11
                            Case "vchUser12"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12
                            Case "vchUser13"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13
                            Case "vchUser14"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14
                            Case "vchUser15"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15
                            Case "vchUser16"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16
                            Case "vchUser17"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17
                            Case "vchUser18"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18
                            Case "vchUser19"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19
                            Case "vchUser20"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20
                            Case "vchUser21"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21
                            Case "vchUser22"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22
                            Case "vchUser23"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23
                            Case "vchUser24"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24
                            Case "vchUser25"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25
                            Case "vchUser26"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26
                            Case "vchUser27"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27
                            Case "vchUser28"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28
                            Case "vchUser29"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29
                            Case "vchUser30"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30
                            Case "vchUser31"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31
                            Case "vchUser32"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32
                            Case "vchUser33"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33
                            Case "vchUser34"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34
                            Case "vchUser35"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35
                            Case "vchUser36"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36
                            Case "vchUser37"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37
                            Case "vchUser38"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38
                            Case "vchUser39"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39
                            Case "vchUser40"
                                varValorLista.proValorID = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40
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
            
            Me.grdDetalles.AddItem varValor
            
            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proRecordStatus = 0 Then
                Me.grdDetalles.RowHeight(Me.grdDetalles.Rows - 1) = 0
            Else
                If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId = "A" Then
                    varCantidadRegistros = varCantidadRegistros + 1
                End If
            End If
            
            'No mostrar los registros cancelados
            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId = "C" And Me.chkCancelados(Index).Value = False Then
                Me.grdDetalles.RowHeight(Me.grdDetalles.Rows - 1) = 0
            Else
                If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId = "C" Then
                    Call SubFPintarFila(Me.grdDetalles, Me.grdDetalles.Rows - 1, Me.lblCancelados(Index).BackColor)
                End If
            End If
            
            'Colorear los registros seleccionados
            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId <> "C" Then
                If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "0" Then
                    Call SubFPintarFila(Me.grdDetalles, varContador, Me.lblSinSeleccion(Index).BackColor)
                Else
                    Call SubFPintarFila(Me.grdDetalles, varContador, Me.lblSeleccion(Index).BackColor)
                End If
            End If
        Next
        Me.grdDetalles.Redraw = True
        Me.txtCantidadRegistros.Text = varCantidadRegistros
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridNumeroPublico(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdNumeroPublico.Rows = 1
    
    For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
        Me.grdNumeroPublico.AddItem Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proDatosProductoId & vbTab & _
                                    Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proTipoLinea & vbTab & _
                                    Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionCode & vbTab & _
                                    Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proRegionName & vbTab & _
                                    Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proNumero & vbTab & _
                                    Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proClasificacionDescripcion & vbTab & _
                                    Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proFechaAsignacion & vbTab & _
                                    Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proPublicar
                                    
        If Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proSeleccion = "0" Then
            Call SubFPintarFila(Me.grdNumeroPublico, varContador, Me.lblSinSeleccion(Index).BackColor)
        Else
            Call SubFPintarFila(Me.grdNumeroPublico, varContador, Me.lblSeleccion(Index).BackColor)
        End If
    Next varContador
    
    Me.grdNumeroPublico.Row = 0
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridTiposLineaModificacion(Index As Integer)
    Dim varContador As Integer
    Dim varContadorAux As Integer
    Dim varValor As String
    Dim varValorCampo As String
    Dim varValorLista As EDCAdminVoz.claValor
    Dim varCantidadRegistros As Integer
    On Error GoTo ErrManager
    
    'Tipos de líneas
    If Index = 0 Then
    
        If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
            Exit Sub
        End If
        
        varValor = ""
        
        Screen.MousePointer = 11
        Me.grdDetallesModificacion.Redraw = False
        Me.grdDetallesModificacion.Rows = 1
        varCantidadRegistros = 0
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
        
            varValor = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId
            varValor = varValor & vbTab & Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proNovedadDetalleDatosProductoId

            For varContadorAux = 1 To Me.proDatosProducto.proParametrosProducto.Count
                Select Case Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proTipo
                    Case "L"
                        Set varValorLista = New EDCAdminVoz.claValor
                        Set varValorLista.proConexion = Me.proConexion
                    
                        Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proCampo)
                            Case "vchUser1"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1
                            Case "vchUser2"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2
                            Case "vchUser3"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3
                            Case "vchUser4"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4
                            Case "vchUser5"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5
                            Case "vchUser6"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6
                            Case "vchUser7"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7
                            Case "vchUser8"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8
                            Case "vchUser9"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9
                            Case "vchUser10"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10
                            Case "vchUser11"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11
                            Case "vchUser12"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12
                            Case "vchUser13"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13
                            Case "vchUser14"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14
                            Case "vchUser15"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15
                            Case "vchUser16"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16
                            Case "vchUser17"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17
                            Case "vchUser18"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18
                            Case "vchUser19"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19
                            Case "vchUser20"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20
                            Case "vchUser21"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21
                            Case "vchUser22"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22
                            Case "vchUser23"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23
                            Case "vchUser24"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24
                            Case "vchUser25"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25
                            Case "vchUser26"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26
                            Case "vchUser27"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27
                            Case "vchUser28"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28
                            Case "vchUser29"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29
                            Case "vchUser30"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30
                            Case "vchUser31"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31
                            Case "vchUser32"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32
                            Case "vchUser33"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33
                            Case "vchUser34"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34
                            Case "vchUser35"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35
                            Case "vchUser36"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36
                            Case "vchUser37"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37
                            Case "vchUser38"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38
                            Case "vchUser39"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39
                            Case "vchUser40"
                                varValorLista.proValorID = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40
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
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser2"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser3"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser4"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser5"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser6"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser7"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser8"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser9"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser10"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser11"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser12"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser13"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser14"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser15"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser16"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser17"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser18"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser19"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser20"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser21"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser22"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser23"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser24"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser25"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser26"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser27"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser28"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser29"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser30"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser31"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser32"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser33"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser34"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser35"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser36"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser37"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser38"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser39"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                            Case "vchUser40"
                                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40 = 1 Then
                                    varValorCampo = "SI"
                                Else
                                    varValorCampo = "NO"
                                End If
                        End Select
                    
                        varValor = varValor & vbTab & varValorCampo
                    Case Else
                    
                        Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proCampo)
                            Case "vchUser1"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1
                            Case "vchUser2"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2
                            Case "vchUser3"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3
                            Case "vchUser4"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4
                            Case "vchUser5"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5
                            Case "vchUser6"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6
                            Case "vchUser7"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7
                            Case "vchUser8"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8
                            Case "vchUser9"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9
                            Case "vchUser10"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10
                            Case "vchUser11"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11
                            Case "vchUser12"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12
                            Case "vchUser13"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13
                            Case "vchUser14"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14
                            Case "vchUser15"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15
                            Case "vchUser16"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16
                            Case "vchUser17"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17
                            Case "vchUser18"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18
                            Case "vchUser19"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19
                            Case "vchUser20"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20
                            Case "vchUser21"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21
                            Case "vchUser22"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22
                            Case "vchUser23"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23
                            Case "vchUser24"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24
                            Case "vchUser25"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25
                            Case "vchUser26"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26
                            Case "vchUser27"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27
                            Case "vchUser28"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28
                            Case "vchUser29"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29
                            Case "vchUser30"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30
                            Case "vchUser31"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31
                            Case "vchUser32"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32
                            Case "vchUser33"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33
                            Case "vchUser34"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34
                            Case "vchUser35"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35
                            Case "vchUser36"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36
                            Case "vchUser37"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37
                            Case "vchUser38"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38
                            Case "vchUser39"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39
                            Case "vchUser40"
                                varValorCampo = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40
                        End Select
                    
                        varValor = varValor & vbTab & varValorCampo
                End Select
            Next varContadorAux
                    
           Me.grdDetallesModificacion.AddItem varValor
            
            'No mostrar registros eliminados
            If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proRecordStatus = 0 Then
                Me.grdDetallesModificacion.RowHeight(Me.grdDetallesModificacion.Rows - 1) = 0
            Else
                varCantidadRegistros = varCantidadRegistros + 1
            End If
            
            If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "0" Then
                Select Case Val(Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proTipoNovedadId)
                    Case 1
                        Call SubFPintarFila(Me.grdDetallesModificacion, Me.grdDetallesModificacion.Rows - 1, Me.lblInsertar(Index).BackColor)
                    Case 2
                        Call SubFPintarFila(Me.grdDetallesModificacion, Me.grdDetallesModificacion.Rows - 1, Me.lblModificar(Index).BackColor)
                    Case 3
                        Call SubFPintarFila(Me.grdDetallesModificacion, Me.grdDetallesModificacion.Rows - 1, Me.lblEliminar(Index).BackColor)
                End Select
            Else
                Call SubFPintarFila(Me.grdDetallesModificacion, Me.grdDetallesModificacion.Rows - 1, Me.lblSeleccionModificacion(Index).BackColor)
            End If
        Next
        
        Me.grdDetallesModificacion.Row = 0
        Me.grdDetallesModificacion.Redraw = True
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub

Private Sub SubFPintarGridEdicionNumeroPublico(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    'Consultar números en proceso de instalación o modificación
    If Me.proDatosProducto.proNovedadNumero Is Nothing Then
        Set Me.proDatosProducto.proNovedadNumero = New colNovedadNumero
        Set Me.proDatosProducto.proNovedadNumero.proConexion = Me.proConexion
        
        If Not Me.proDatosProducto.MetConsultarNovedadNumeros Then
            MsgBox "Error al consultar los números  que se encuentran en proceso de instalación o modificación.", vbCritical, App.Title
            Exit Sub
        End If
    End If
            
    Me.grdEdicionNumeroPublico.Rows = 1
    
    For varContador = 1 To Me.proDatosProducto.proNovedadNumero.Count
        Me.grdEdicionNumeroPublico.AddItem Me.proDatosProducto.proNovedadNumero.Item(varContador).proNovedadNumeroId & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proTipoLinea & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proRegionCode & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proRegionName & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proNumero & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proClasificacionDescripcion & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proDatosProductoId & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proIncidentId & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proTipoNovedadId & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proFechaReserva & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proFechaLiberacion & vbTab & _
                                           Me.proDatosProducto.proNovedadNumero.Item(varContador).proPublicar
                                           
        If Me.proDatosProducto.proNovedadNumero.Item(varContador).proSeleccion = "0" Then
            Select Case Val(Me.proDatosProducto.proNovedadNumero.Item(varContador).proTipoNovedadId)
                Case 1
                    Call SubFPintarFila(Me.grdEdicionNumeroPublico, Me.grdEdicionNumeroPublico.Rows - 1, Me.lblInsertar(Index).BackColor)
                Case 2
                    Call SubFPintarFila(Me.grdEdicionNumeroPublico, Me.grdEdicionNumeroPublico.Rows - 1, Me.lblModificar(Index).BackColor)
                Case 3
                    Call SubFPintarFila(Me.grdEdicionNumeroPublico, Me.grdEdicionNumeroPublico.Rows - 1, Me.lblEliminar(Index).BackColor)
            End Select
        Else
            Call SubFPintarFila(Me.grdEdicionNumeroPublico, Me.grdEdicionNumeroPublico.Rows - 1, Me.lblSeleccionModificacion(Index).BackColor)
        End If
        
    Next varContador
    
    Me.grdEdicionNumeroPublico.Row = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridEdicionNumeracionCorporativa(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdEdicionNumeracionPrivada.Rows = 1
    
    For varContador = 1 To Me.proDatosProducto.proNovedadNumeracionCorporativa.Count
        Me.grdEdicionNumeracionPrivada.AddItem Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proDatosProductoId & vbTab & _
                                           Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proIncidentId & vbTab & _
                                           Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proTipoNovedadId & vbTab & _
                                           Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proMarcacion & vbTab & _
                                           Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proVirtual
                                           'Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proVirtual  Agregado por Carlos Castelblanco 2006/07/26
                                           
        If Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proSeleccion = "0" Then
            Select Case Val(Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proTipoNovedadId)
                Case 1
                    Call SubFPintarFila(Me.grdEdicionNumeracionPrivada, Me.grdEdicionNumeracionPrivada.Rows - 1, Me.lblInsertar(Index).BackColor)
                Case 2
                    Call SubFPintarFila(Me.grdEdicionNumeracionPrivada, Me.grdEdicionNumeracionPrivada.Rows - 1, Me.lblModificar(Index).BackColor)
                Case 3
                    Call SubFPintarFila(Me.grdEdicionNumeracionPrivada, Me.grdEdicionNumeracionPrivada.Rows - 1, Me.lblEliminar(Index).BackColor)
            End Select
        Else
            Call SubFPintarFila(Me.grdEdicionNumeracionPrivada, Me.grdEdicionNumeracionPrivada.Rows - 1, Me.lblSeleccionModificacion(Index).BackColor)
        End If
        
    Next varContador
    
    Me.grdEdicionNumeracionPrivada.Row = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridServiciosxNumero(Index As Integer)
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
                varServiciosRelacionados = "[" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNombreServicio & "],"
            Else
                If varRegionCodeAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proRegionCode _
                   And varNumeroAnterior = Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNumero Then
                   
                    varServiciosRelacionados = varServiciosRelacionados & " [" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNombreServicio & "],"
                   
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
                    
                    varServiciosRelacionados = "[" & Me.proDatosProducto.proServiciosxNumero.Item(varContador).proNombreServicio & "],"
                    
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
Private Sub SubFInicializarGridTiposLinea(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
        
    'Tipos de línea
        If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
            Me.grdDetalles.Rows = 0
            Me.cmdEliminar(Index).Enabled = False
            Me.cmdInsertar(Index).Enabled = False
            Me.cmdModificar(Index).Enabled = False
            Me.cmdModificarColumna(Index).Enabled = False
            Me.cmdClonar(Index).Enabled = False
            MsgBox "El producto del incidente seleccionado no tiene campos parametrizados."
            Exit Sub
        Else
            If varInsertarTiposLinea And Me.proDatosProducto.proiEstratoid <> "" Then
                Me.cmdInsertar(Index).Enabled = True
            Else
                Me.cmdInsertar(Index).Enabled = False
            End If
            
            Me.cmdEliminar(Index).Enabled = False
            Me.cmdModificar(Index).Enabled = False
            Me.cmdModificarColumna(Index).Enabled = False
            Me.cmdClonar(Index).Enabled = False
        End If
        
        With Me.grdDetalles
            .Cols = Me.proDatosProducto.proParametrosProducto.Count + 2
            .Rows = 1
            .Row = 0
            
            .Col = 0
            .CellAlignment = 4
            .ColWidth(0) = 800
            .TextMatrix(0, 0) = "# Línea"
            
            .Col = 1
            .CellAlignment = 4
            .ColWidth(1) = 0
            .TextMatrix(0, 1) = "" 'Codigo Estado
                
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
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFInicializarGridNumeroPublico(Index As Integer)
    On Error GoTo ErrManager
    
    With Me.grdNumeroPublico
        .Rows = 1
        .Cols = 8
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 0
        .TextMatrix(0, 0) = "" 'DatosProductoId
         
        .Col = 1
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(1) = 800
        .TextMatrix(0, 1) = "# Línea"
       
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 0
        .TextMatrix(0, 2) = "" 'Codigo Ciudad
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 1100
        .TextMatrix(0, 3) = "Ciudad"
        
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "Número"
        
        'Columna agregada por Carlos Castelblanco 2006/07/28:
        .Col = 5
        .CellAlignment = 4
        .ColWidth(5) = 2050
        .TextMatrix(0, 5) = "Clasificacion"
        
        'Columna modificada por Carlos Castelblanco 2006/07/28:
        .Col = 6
        .CellAlignment = 4
        .ColWidth(6) = 2050
        .TextMatrix(0, 6) = "Fecha Asignación"
        
        .Col = 7
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(7) = 2300
        .TextMatrix(0, 7) = "Publicar (S/N)"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridNumeroCorporativo(Index As Integer)
    On Error GoTo ErrManager
    
    With Me.grdNumeracionPrivada
        .Rows = 1
        .Cols = 3
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 1000
        .TextMatrix(0, 0) = "DatosProductoId"
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 2000
        .TextMatrix(0, 1) = "Marcación"
        
        'Agregado por Carlos Castelblanco 2006/07/26:
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = "Virtual"
        
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Public Sub SubFInicializarGridNumeroCorporativoModificados(Index As Integer)
    On Error GoTo ErrManager
    
    With Me.grdEdicionNumeracionPrivada
        .Rows = 1
        .Cols = 5 'Modificado por Carlos Castelblanco 2006/07/28
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 1000
        .TextMatrix(0, 0) = "DatosProductoId"
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 1000
        .TextMatrix(0, 1) = "Incidente"
        
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = "Tipo Novedad"
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = "Marcacion"
        
        'Agregado por Carlos Castelblanco 2006/07/26:
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Virtual"
        
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFInicializarGridNumeroPublicoModificados(Index As Integer)
    On Error GoTo ErrManager
    
    With Me.grdEdicionNumeroPublico
        .Rows = 1
        .Cols = 12
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 0
        .TextMatrix(0, 0) = "" 'NovedadNumeroId
        
        .Col = 1
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(1) = 800
        .TextMatrix(0, 1) = "# Línea"
        
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 0
        .TextMatrix(0, 2) = "" 'Codigo Ciudad
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = "Ciudad"
        
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 1500
        .TextMatrix(0, 4) = "Número"
        
        'Columna Agregada por Carlos Castelblanco 2006/07/28:
        .Col = 5
        .CellAlignment = 4
        .ColWidth(5) = 2050
        .TextMatrix(0, 5) = "Clasificacion"
                
        'Indice incrementado en 1 por Carlos Castelblanco 2006/07/28:
        .Col = 6
        .CellAlignment = 4
        .ColWidth(6) = 0
        .TextMatrix(0, 6) = "DatosProductoId"
        
        'Indice incrementado en 1 por Carlos Castelblanco 2006/07/28:
        .Col = 7
        .CellAlignment = 4
        .ColWidth(7) = 1500
        .TextMatrix(0, 7) = "Incidente"
        
        'Indice incrementado en 1 por Carlos Castelblanco 2006/07/28:
        .Col = 8
        .CellAlignment = 4
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = "Tipo Novedad"
        
        'Indice incrementado en 1 por Carlos Castelblanco 2006/07/28:
        .Col = 9
        .CellAlignment = 4
        .ColWidth(9) = 2000
        .TextMatrix(0, 9) = "Fecha de Reserva"
        
        'Indice incrementado en 1 por Carlos Castelblanco 2006/07/28:
        .Col = 10
        .CellAlignment = 4
        .ColWidth(10) = 2000
        .TextMatrix(0, 10) = "Fecha de Liberación"
        
        .Col = 11
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(11) = 1500
        .TextMatrix(0, 11) = "Publicar (S/N)"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFInicializarGridTiposLineaModificados(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
        
    'Tipos de líneas
    If Index = 0 Then
        If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
            Me.grdDetallesModificacion.Rows = 0
            Me.cmdEliminar(Index).Enabled = False
            Me.cmdInsertar(Index).Enabled = False
            Me.cmdModificar(Index).Enabled = False
            Me.cmdModificarColumna(Index).Enabled = False
            Me.cmdClonar(Index).Enabled = False
            MsgBox "El producto del incidente seleccionado no tiene campos parametrizados."
            Exit Sub
        Else
            If varInsertarTiposLinea And Me.proDatosProducto.proiEstratoid <> "" Then
                Me.cmdInsertar(Index).Enabled = True
            Else
                Me.cmdInsertar(Index).Enabled = False
            End If
            
            Me.cmdEliminar(Index).Enabled = False
            Me.cmdModificar(Index).Enabled = False
            Me.cmdModificarColumna(Index).Enabled = False
            Me.cmdClonar(Index).Enabled = False
        End If
        
        With Me.grdDetallesModificacion
            .Cols = Me.proDatosProducto.proParametrosProducto.Count + 2
            .Rows = 1
            .Row = 0
            
            .Col = 0
            .CellAlignment = 4
            .ColWidth(0) = 0
            .TextMatrix(0, 0) = "" 'Código
            
            .Col = 1
            .CellAlignment = 4
            .ColWidth(1) = 800
            .TextMatrix(0, 1) = "# Línea"
            
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

Private Sub SubFInicializarGridServiciosSuplementarios(Index As Integer)
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
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = "Ciudad"
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 1100
        .TextMatrix(0, 3) = "Número"
        
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 3500
        .TextMatrix(0, 4) = "Servicios Asignados"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim varContador As Integer, varIndice As Integer
    On Error GoTo ErrManager
    If Trim(Me.proDatosProducto.proTipoTelefonia) = "107441" Then
        Exit Sub
    End If
    'Verificar que los tipos de línea en edición tengan el número mínimo de lineas asignadas
    For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
        If proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proTipoNovedadId <> "2" And _
            proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proTipoNovedadId <> "3" And _
            proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15 <> varSi Then 'No validar novedad de modificación, eliminación o backups
            varIndice = varValoresCampoProductoTipoLinea.BuscarIndiceProValorId(Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1)
            If varIndice > -1 Then
                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proContadorNumeros < varValoresCampoProductoTipoLinea.Item(varIndice).proMinimo Then
                    MsgBox "El tipo de línea en edición con número de línea " & proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proNovedadDetalleDatosProductoId & " no tiene la cantidad mínima de números asignados (" & varValoresCampoProductoTipoLinea.Item(varIndice).proMinimo & ")", vbCritical, App.Title
                   ' Cancel = 1
                    'Exit Sub
                End If
            End If
        End If
    Next
    
    If Not Me.proDatosProducto.proDetalleDatosProducto Is Nothing Then
        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = 0
        Next varContador
        
        Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 0
    End If
    
    If Not Me.proDatosProducto.proNumeracionCorporativa Is Nothing Then
        For varContador = 1 To Me.proDatosProducto.proNumeracionCorporativa.Count
            Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proSeleccion = 0
        Next varContador
        
        Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados = 0
    End If
    
    If Not Me.proDatosProducto.proNovedadDetalleDatosProducto Is Nothing Then
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = 0
        Next varContador
        
        Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 0
    End If
    
    If Not Me.proDatosProducto.proDatosProductoNumero Is Nothing Then
        For varContador = 1 To Me.proDatosProducto.proDatosProductoNumero.Count
            Me.proDatosProducto.proDatosProductoNumero.Item(varContador).proSeleccion = 0
        Next varContador
        
        Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = 0
    End If
    
    If Not Me.proDatosProducto.proNovedadNumeracionCorporativa Is Nothing Then
        
        For varContador = 1 To Me.proDatosProducto.proNovedadNumeracionCorporativa.Count
            Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proSeleccion = 0
        Next varContador
        
        Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados = 0
    End If
    
    If Not Me.proDatosProducto.proNovedadNumero Is Nothing Then
        For varContador = 1 To Me.proDatosProducto.proNovedadNumero.Count
            Me.proDatosProducto.proNovedadNumero.Item(varContador).proSeleccion = 0
        Next varContador
        
        Me.proDatosProducto.proNovedadNumero.proSeleccionados = 0
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdDetalles_DblClick()
    Dim Index As Integer
    On Error GoTo ErrManager
    
    Index = Me.TbFondo.Tab
    
    'Tipos de líneas
    If Index = 0 Then
    
        If Me.grdDetalles.Row = 0 Then
            Exit Sub
        End If
        
        If Me.proDatosProducto.proDetalleDatosProducto.Item(Me.grdDetalles.Row).proStatusId = "A" Then
            If Me.proDatosProducto.proDetalleDatosProducto.Item(Me.grdDetalles.Row).proSeleccion = "0" Then
                Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados + 1
                Me.proDatosProducto.proDetalleDatosProducto.Item(Me.grdDetalles.Row).proSeleccion = 1
                Call SubFPintarFila(Me.grdDetalles, Me.grdDetalles.Row, Me.lblSeleccion(Index).BackColor)
            Else
                Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados - 1
                Me.proDatosProducto.proDetalleDatosProducto.Item(Me.grdDetalles.Row).proSeleccion = 0
                Call SubFPintarFila(Me.grdDetalles, Me.grdDetalles.Row, Me.lblSinSeleccion(Index).BackColor)
            End If
        End If
        
        If Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados <> 0 Then
            If varModificarTiposLinea Then
                Me.cmdModificar(Index).Enabled = True
                Me.cmdModificarColumna(Index).Enabled = True
            End If
            
            If varEliminarTiposLinea Then
                Me.cmdEliminar(Index).Enabled = True
            End If
            
            If varInsertarTiposLinea Then
                Me.cmdClonar(Index).Enabled = True
            End If
        Else
            Me.cmdModificar(Index).Enabled = False
            Me.cmdModificarColumna(Index).Enabled = False
            Me.cmdEliminar(Index).Enabled = False
            Me.cmdClonar(Index).Enabled = False
        End If
        'MsgBox Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdDetallesModificacion_DblClick()
    Dim Index As Integer
    On Error GoTo ErrManager
    
    Index = Me.TbFondo.Tab
    
    'Tipos de líneas
    If Index = 0 Then
    
        If Me.grdDetallesModificacion.Row = 0 Then
            Exit Sub
        End If
        
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.grdDetallesModificacion.Row).proSeleccion = "0" Then
            Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados + 1
            Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.grdDetallesModificacion.Row).proSeleccion = 1
            Call SubFPintarFila(Me.grdDetallesModificacion, Me.grdDetallesModificacion.Row, Me.lblSeleccionModificacion(Index).BackColor)
        Else
            Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados - 1
            Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.grdDetallesModificacion.Row).proSeleccion = 0
            Select Case Val(Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.grdDetallesModificacion.Row).proTipoNovedadId)
                Case 1
                    Call SubFPintarFila(Me.grdDetallesModificacion, Me.grdDetallesModificacion.Row, Me.lblInsertar(Index).BackColor)
                Case 2
                    Call SubFPintarFila(Me.grdDetallesModificacion, Me.grdDetallesModificacion.Row, Me.lblModificar(Index).BackColor)
                Case 3
                    Call SubFPintarFila(Me.grdDetallesModificacion, Me.grdDetallesModificacion.Row, Me.lblEliminar(Index).BackColor)
            End Select
        End If
        
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados <> 0 Then
            If varInsertarTiposLinea Then
                Me.cmdClonarModificados(Index).Enabled = True
            Else
                Me.cmdClonarModificados(Index).Enabled = False
            End If
            
            If varModificarTiposLinea Then
                Me.cmdModificarInsertados(Index).Enabled = True
                Me.cmdModificarColumnaInsertados(Index).Enabled = True
            Else
                Me.cmdModificarInsertados(Index).Enabled = False
                Me.cmdModificarColumnaInsertados(Index).Enabled = False
            End If
    
            If varInsertarTiposLinea Or varModificarTiposLinea Or varEliminarTiposLinea Then
                Me.cmdDeshacerModificación(Index).Enabled = True
            Else
                Me.cmdDeshacerModificación(Index).Enabled = False
            End If
        Else
            Me.cmdDeshacerModificación(Index).Enabled = False
            Me.cmdClonarModificados(Index).Enabled = False
            Me.cmdModificarInsertados(Index).Enabled = False
            Me.cmdModificarColumnaInsertados(Index).Enabled = False
        End If
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub



Private Sub grdEdicionNumeracionPrivada_DblClick()
    Dim Index As Integer
    On Error GoTo ErrManager
    
    Index = Me.TbFondo.Tab
    
    If Me.grdEdicionNumeracionPrivada.Row = 0 Then
        Exit Sub
    End If
    
    If Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(Me.grdEdicionNumeracionPrivada.Row).proSeleccion = "0" Then
        Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados = Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados + 1
        Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(Me.grdEdicionNumeracionPrivada.Row).proSeleccion = 1
        Call SubFPintarFila(Me.grdEdicionNumeracionPrivada, Me.grdEdicionNumeracionPrivada.Row, Me.lblSeleccionModificacion(Index).BackColor)
    Else
        Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados = Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados - 1
        Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(Me.grdEdicionNumeracionPrivada.Row).proSeleccion = 0
        Select Case Val(Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(Me.grdEdicionNumeracionPrivada.Row).proTipoNovedadId)
            Case 1
                Call SubFPintarFila(Me.grdEdicionNumeracionPrivada, Me.grdEdicionNumeracionPrivada.Row, Me.lblInsertar(Index).BackColor)
            Case 2
                Call SubFPintarFila(Me.grdEdicionNumeracionPrivada, Me.grdEdicionNumeracionPrivada.Row, Me.lblModificar(Index).BackColor)
            Case 3
                Call SubFPintarFila(Me.grdEdicionNumeracionPrivada, Me.grdEdicionNumeracionPrivada.Row, Me.lblEliminar(Index).BackColor)
        End Select
    End If
    
    If Me.proDatosProducto.proNovedadNumeracionCorporativa.proSeleccionados <> 0 Then

        If varInsertarNumeracionCorporativa Or varModificarNumeracionCorporativa Or varEliminarNumeracionCorporativa Then
            Me.cmdDeshacerModificación(Index).Enabled = True
        Else
            Me.cmdDeshacerModificación(Index).Enabled = False
        End If
    Else
        Me.cmdDeshacerModificación(Index).Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdEdicionNumeroPublico_DblClick()
    Dim Index As Integer
    On Error GoTo ErrManager
    
    Index = Me.TbFondo.Tab
    
    If Me.grdEdicionNumeroPublico.Row = 0 Then
        Exit Sub
    End If
    
    If Me.proDatosProducto.proNovedadNumero.Item(Me.grdEdicionNumeroPublico.Row).proSeleccion = "0" Then
        Me.proDatosProducto.proNovedadNumero.proSeleccionados = Me.proDatosProducto.proNovedadNumero.proSeleccionados + 1
        Me.proDatosProducto.proNovedadNumero.Item(Me.grdEdicionNumeroPublico.Row).proSeleccion = 1
        Call SubFPintarFila(Me.grdEdicionNumeroPublico, Me.grdEdicionNumeroPublico.Row, Me.lblSeleccionModificacion(Index).BackColor)
    Else
        Me.proDatosProducto.proNovedadNumero.proSeleccionados = Me.proDatosProducto.proNovedadNumero.proSeleccionados - 1
        Me.proDatosProducto.proNovedadNumero.Item(Me.grdEdicionNumeroPublico.Row).proSeleccion = 0
        Select Case Val(Me.proDatosProducto.proNovedadNumero.Item(Me.grdEdicionNumeroPublico.Row).proTipoNovedadId)
            Case 1
                Call SubFPintarFila(Me.grdEdicionNumeroPublico, Me.grdEdicionNumeroPublico.Row, Me.lblInsertar(Index).BackColor)
            Case 2
                Call SubFPintarFila(Me.grdEdicionNumeroPublico, Me.grdEdicionNumeroPublico.Row, Me.lblModificar(Index).BackColor)
            Case 3
                Call SubFPintarFila(Me.grdEdicionNumeroPublico, Me.grdEdicionNumeroPublico.Row, Me.lblEliminar(Index).BackColor)
        End Select
    End If
    
    If Me.proDatosProducto.proNovedadNumero.proSeleccionados <> 0 Then

        If varInsertarNumeracionPublica Or varModificarNumeracionPublica Or varEliminarNumeracionPublica Then
            Me.cmdDeshacerModificación(Index).Enabled = True
        Else
            Me.cmdDeshacerModificación(Index).Enabled = False
        End If
    Else
        Me.cmdDeshacerModificación(Index).Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdNumeracionPrivada_DblClick()
    Dim Index As Integer
    On Error GoTo ErrManager
    
    Index = Me.TbFondo.Tab
    
    If Me.grdNumeracionPrivada.Row = 0 Then
        Exit Sub
    End If
    
    If Me.proDatosProducto.proNumeracionCorporativa.Item(Me.grdNumeracionPrivada.Row).proSeleccion = "0" Then
        Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados = Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados + 1
        Me.proDatosProducto.proNumeracionCorporativa.Item(Me.grdNumeracionPrivada.Row).proSeleccion = 1
        Call SubFPintarFila(Me.grdNumeracionPrivada, Me.grdNumeracionPrivada.Row, Me.lblSeleccionModificacion(Index).BackColor)
    Else
        Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados = Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados - 1
        Me.proDatosProducto.proNumeracionCorporativa.Item(Me.grdNumeracionPrivada.Row).proSeleccion = 0
        
        Call SubFPintarFila(Me.grdNumeracionPrivada, Me.grdNumeracionPrivada.Row, Me.lblSinSeleccion(Index).BackColor)
    End If
    
    If Me.proDatosProducto.proNumeracionCorporativa.proSeleccionados <> 0 Then
        
        If varInsertarNumeracionCorporativa Or varModificarNumeracionCorporativa Or varEliminarNumeracionCorporativa Then
            Me.cmdEliminar(Index).Enabled = True
        Else
            Me.cmdEliminar(Index).Enabled = False
        End If
    Else
        Me.cmdEliminar(Index).Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError

End Sub

Private Sub grdNumeroPublico_DblClick()
    Dim Index As Integer
    On Error GoTo ErrManager
    
    Index = Me.TbFondo.Tab
    
    If Me.grdNumeroPublico.Row = 0 Then
        Exit Sub
    End If
    
    If Me.proDatosProducto.proDatosProductoNumero.Item(Me.grdNumeroPublico.Row).proSeleccion = "0" Then
        Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = Me.proDatosProducto.proDatosProductoNumero.proSeleccionados + 1
        Me.proDatosProducto.proDatosProductoNumero.Item(Me.grdNumeroPublico.Row).proSeleccion = 1
        Call SubFPintarFila(Me.grdNumeroPublico, Me.grdNumeroPublico.Row, Me.lblSeleccionModificacion(Index).BackColor)
    Else
        Me.proDatosProducto.proDatosProductoNumero.proSeleccionados = Me.proDatosProducto.proDatosProductoNumero.proSeleccionados - 1
        Me.proDatosProducto.proDatosProductoNumero.Item(Me.grdNumeroPublico.Row).proSeleccion = 0
        
        Call SubFPintarFila(Me.grdNumeroPublico, Me.grdNumeroPublico.Row, Me.lblSinSeleccion(Index).BackColor)
    End If
    
    If Me.proDatosProducto.proDatosProductoNumero.proSeleccionados <> 0 Then
            Me.CmdCambiarTipoLinea.Enabled = True '-->3.7.4
        If varInsertarNumeracionPublica Or varModificarNumeracionPublica Or varEliminarNumeracionPublica Then
            Me.cmdEliminar(Index).Enabled = True
        Else
            Me.cmdEliminar(Index).Enabled = False
        End If
    Else
        Me.cmdEliminar(Index).Enabled = False
        Me.CmdCambiarTipoLinea.Enabled = False '-->3.7.4
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub




Private Sub lstGrupoCentrexCallSource_LostFocus()
    On Error GoTo ErrManager
    
    Me.pnlGrupoCentrexCallSource.Visible = False
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub TbFondo_Click(PreviousTab As Integer)
    Dim Index As Integer
    On Error GoTo ErrManager
    
    Index = Me.TbFondo.Tab
        
    Me.txtCodigoProducto(Index).Text = Me.proDatosProducto.proProductNumber
    Me.txtNombreProducto(Index).Text = Me.proDatosProducto.proProductName
    Me.txtIncidente(Index).Text = Me.proDatosProducto.proIncidentId
            
    'Tipos de Líneas
    If Index = 0 Then
        If Not varTiposLinea Then
            varTiposLinea = True
            
            Call SubFInicializarGridTiposLinea(Index)
            Call SubFInicializarGridTiposLineaModificados(Index)
            
            If Me.proDatosProducto.proDetalleDatosProducto Is Nothing Then
                Set Me.proDatosProducto.proDetalleDatosProducto = New colDetalleDatosProducto
                Set Me.proDatosProducto.proDetalleDatosProducto.proConexion = Me.proConexion
                
                If Me.proDatosProducto.MetConsultarDetalles Then
                    Call SubFPintarGridTiposLinea(Index)
                Else
                    MsgBox "Error al recuperar la información de los detalles.", vbCritical, App.Title
                    Exit Sub
                End If
            Else
                Call SubFPintarGridTiposLinea(Index)
            End If
              
            If Me.proDatosProducto.MetConsultarNovedadDetalleDatosProducto Then
                Call SubFPintarGridTiposLineaModificacion(Index)
            Else
                MsgBox "Error al consultar los registros en curso de modificación.", vbCritical, App.Title
            End If
            
            If varInsertarTiposLinea Or varModificarTiposLinea Then
                Me.cmdBuscarCliente.Enabled = True
                Me.cmdEliminarCliente.Enabled = True
            Else
                Me.cmdBuscarCliente.Enabled = False
                Me.cmdEliminarCliente.Enabled = False
            End If
        End If
    End If
    
    'Numeración Privada
    If Index = 1 Then
        If Not varNumerosPrivados Then
            varNumerosPrivados = True
            
            Call SubFInicializarGridNumeroCorporativo(Index)
            Call SubFInicializarGridNumeroCorporativoModificados(Index)
            
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
            
            'Consultar del Plan de Numeración En Curso
            Set varPlanNumeracionEnCurso = New colPlanNumeracionEnCurso
            Set varPlanNumeracionEnCurso.proConexion = Me.proConexion
            varPlanNumeracionEnCurso.proCliente = Me.proOnyx.ContactID
            
            If varPlanNumeracionEnCurso.MetConsultaEnCurso Then
                Call SubFPintarTreePlanNumeracionEnCurso
            Else
                MsgBox "Error al consultar el plan de numeración en curso de instalación.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Consultar los numeros instalados actualmente
            If Me.proDatosProducto.MetConsultarNumeracionCorporativa Then
                Call SubFPintarGridNumeracionCorporativa(Index)
            Else
                MsgBox "Error al Consultar la numeración corporativa.", vbCritical, App.Title
                Exit Sub
            End If
            
            'Consultar los numeros es curso de instalacion
            If Me.proDatosProducto.MetConsultarNovedadNumeracionCorporativa Then
                Call SubFPintarGridEdicionNumeracionCorporativa(Index)
            Else
                MsgBox "Error al consultar la numeración corporativa en curso de moficación.", vbCritical, App.Title
                Exit Sub
            End If
            
            If varInsertarNumeracionCorporativa Then
                Me.cmdInsertar.Item(Index).Enabled = True
            End If
            
        End If
    End If
    
    'Numeración pública
    If Index = 2 Then
        If Not varNumerosPublicos Then
            varNumerosPublicos = True
            
            Call SubFInicializarGridNumeroPublico(Index)
            Call SubFInicializarGridNumeroPublicoModificados(Index)
            Call SubFInicializarGridServiciosSuplementarios(Index)
            
            'Consultar números instalados en el cliente
            If Me.proDatosProducto.proDatosProductoNumero Is Nothing Then
                Set Me.proDatosProducto.proDatosProductoNumero = New colDatosProductoNumero
                Set Me.proDatosProducto.proDatosProductoNumero.proConexion = Me.proConexion
                
                If Me.proDatosProducto.MetConsultarDatosProductoNumero Then
                    Call SubFPintarGridNumeroPublico(Index)
                Else
                    MsgBox "Error al consultar los números asignados al servicio.", vbCritical, App.Title
                    Exit Sub
                End If
            Else
                Call SubFPintarGridNumeroPublico(Index)
            End If
            
            'Consultar números en proceso de instalación o modificación
            If Me.proDatosProducto.proNovedadNumero Is Nothing Then
                Set Me.proDatosProducto.proNovedadNumero = New colNovedadNumero
                Set Me.proDatosProducto.proNovedadNumero.proConexion = Me.proConexion
                
                If Me.proDatosProducto.MetConsultarNovedadNumeros Then
                    Call SubFPintarGridEdicionNumeroPublico(Index)
                Else
                    MsgBox "Error al consultar los números  que se encuentran en proceso de instalación o modificación.", vbCritical, App.Title
                    Exit Sub
                End If
            Else
                Call SubFPintarGridEdicionNumeroPublico(Index)
            End If
            
            'Consulta de servicios suplementarios
            If Me.proDatosProducto.proServiciosxNumero Is Nothing Then
                Set Me.proDatosProducto.proServiciosxNumero = New colServiciosxNumero
                Set Me.proDatosProducto.proServiciosxNumero.proConexion = Me.proConexion
                
                If Me.proDatosProducto.MetConsultarServiciosxNumero Then
                    Call SubFPintarGridServiciosxNumero(Index)
                Else
                    MsgBox "Error al consultar los servicios suplementarios de cada número.", vbCritical, App.Title
                    Exit Sub
                End If
            Else
                Call SubFPintarGridServiciosxNumero(Index)
            End If
            
            If varInsertarNumeracionPublica Then
                Me.cmdInsertar.Item(Index).Enabled = True
            End If
            
            If (Me.proDatosProducto.proNovedadNumero.Count + Me.proDatosProducto.proDatosProductoNumero.Count) <> 0 Then
                If varModificarNumeracionPublica = True Then
                    Me.cmdModificarInsertados(Index).Enabled = True
                Else
                    Me.cmdModificarInsertados(Index).Enabled = False
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridNumeracionCorporativa(Index As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdNumeracionPrivada.Rows = 1
    For varContador = 1 To Me.proDatosProducto.proNumeracionCorporativa.Count
        Me.grdNumeracionPrivada.AddItem Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proDatosProductoId & vbTab & _
                                        Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proMarcacion & vbTab & _
                                        Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proVirtual
                                        'Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proVirtual Agregado por Carlos CAstelblanco 2006/07/26
                                        
                                                   
        If Me.proDatosProducto.proNumeracionCorporativa.Item(varContador).proSeleccion = "0" Then
            Call SubFPintarFila(Me.grdNumeracionPrivada, varContador, Me.lblSinSeleccion(Index).BackColor)
        Else
            Call SubFPintarFila(Me.grdNumeracionPrivada, varContador, Me.lblSeleccion(Index).BackColor)
        End If

    Next varContador
    
    Me.grdNumeracionPrivada.Col = 0
    Me.grdNumeracionPrivada.Row = 0
    
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

Private Sub SubFPintarTreePlanNumeracionEnCurso()
    Dim Nodo As Node
    Dim varContador As Integer
    Dim varIncidenteAnterior As String
    Dim varVirtual As String 'Agregado por Carlos Castelblanco 2006/07/26
    
    On Error GoTo ErrManager
    
    varIncidenteAnterior = ""
    
    Me.trvPlanEnCurso.Nodes.Clear
    
    For varContador = 1 To varPlanNumeracionEnCurso.Count
        'If-Else-End If Agregado por Carlos Castelblanco 2006/07/26
        If varPlanNumeracionEnCurso.Item(varContador).proVirtual = "S" Then
            varVirtual = " - Virtual"
        Else
            varVirtual = " - No Virtual"
        End If
                
        If varContador = 1 Then
            varIncidenteAnterior = varPlanNumeracionEnCurso.Item(varContador).proIncidentId
            
            Set Nodo = Me.trvPlanEnCurso.Nodes.Add(, , "[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "]", "[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "]")
            Set Nodo = Me.trvPlanEnCurso.Nodes.Add("[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "]", tvwChild, "[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "] - [" & varPlanNumeracionEnCurso.Item(varContador).proMarcacion & "]", varPlanNumeracionEnCurso.Item(varContador).proMarcacion & varVirtual)
        Else
            If varIncidenteAnterior = varPlanNumeracionEnCurso.Item(varContador).proIncidentId Then
                Set Nodo = Me.trvPlanEnCurso.Nodes.Add("[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "]", tvwChild, "[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "] - [" & varPlanNumeracionEnCurso.Item(varContador).proMarcacion & "]", varPlanNumeracionEnCurso.Item(varContador).proMarcacion & varVirtual)
            Else
                varIncidenteAnterior = varPlanNumeracionEnCurso.Item(varContador).proIncidentId
                Set Nodo = Me.trvPlanEnCurso.Nodes.Add(, , "[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "]", "[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "]")
                Set Nodo = Me.trvPlanEnCurso.Nodes.Add("[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "]", tvwChild, "[" & varPlanNumeracionEnCurso.Item(varContador).proCategoria & "] - [" & varIncidenteAnterior & "] - [" & varPlanNumeracionEnCurso.Item(varContador).proMarcacion & "]", varPlanNumeracionEnCurso.Item(varContador).proMarcacion & varVirtual)
            End If
        End If
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtCallSource_GotFocus()
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    Me.txtCallSource.SelStart = 0
    Me.txtCallSource.SelLength = Len(Me.txtCallSource.Text)
    
    'Consultar la información del los grupos centrex usados actualmente
    Set varColClienteTelefonia = Nothing
    Set varColClienteTelefonia = New EDCVoz.colClienteTelefonia
    Set varColClienteTelefonia.proConexion = Me.proConexion
    
    If varColClienteTelefonia.MetConsultarCallSourceOcupado Then
        Call SubFPintarCallSource
        Me.txtGrupoCentrexCallSource.Text = "Call Source que se encuentran en uso actualmente."
        Me.pnlGrupoCentrexCallSource.Visible = True
    Else
        MsgBox "Error al consultar los Call Source que se encuentran ocupados actualmente.", vbCritical, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub SubFPintarCallSource()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.lstGrupoCentrexCallSource.Clear
    
    For varContador = 1 To varColClienteTelefonia.Count
        Me.lstGrupoCentrexCallSource.AddItem varColClienteTelefonia.Item(varContador).proCallSource
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtCallSource_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtCallSource_LostFocus()
    On Error GoTo ErrManager
    
    If Me.ActiveControl.Name <> "lstGrupoCentrexCallSource" Then
        Me.pnlGrupoCentrexCallSource.Visible = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub TxtEnlace_LostFocus()
On Error GoTo ErrManager
    Me.TxtIdVenta.Text = ""
    Me.TxtIdVenta.Text = Trim(Me.proDatosProducto.MetDevolverVenta(Trim(TxtEnlace.Text)))
Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtGrupoCentrex_GotFocus()
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    Me.txtGrupoCentrex.SelStart = 0
    Me.txtGrupoCentrex.SelLength = Len(Me.txtGrupoCentrex.Text)
    
    'Consultar la información del los grupos centrex usados actualmente
    Set varColClienteTelefonia = Nothing
    Set varColClienteTelefonia = New EDCVoz.colClienteTelefonia
    Set varColClienteTelefonia.proConexion = Me.proConexion
    
    If varColClienteTelefonia.MetConsultarGrupoCentrexOcupado Then
        Call SubFPintarGrupoCentrex
        Me.txtGrupoCentrexCallSource.Text = "Grupo Centrex que se encuentran en uso actualmente."
        Me.pnlGrupoCentrexCallSource.Visible = True
    Else
        MsgBox "Error al consultar los Grupo Centrex que se encuentran ocupados actualmente.", vbCritical, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub SubFPintarGrupoCentrex()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.lstGrupoCentrexCallSource.Clear
    
    For varContador = 1 To varColClienteTelefonia.Count
        Me.lstGrupoCentrexCallSource.AddItem varColClienteTelefonia.Item(varContador).proGrupoCentrex
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtGrupoCentrex_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtGrupoCentrex_LostFocus()
    On Error GoTo ErrManager
    
    If Me.ActiveControl.Name <> "lstGrupoCentrexCallSource" Then
        Me.pnlGrupoCentrexCallSource.Visible = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub TxtIdVenta_KeyPress(KeyAscii As Integer)
 On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub TxtIdVenta_LostFocus()
On Error GoTo ErrManager
      Me.TxtEnlace.Text = ""
      Me.TxtEnlace.Text = Trim(Me.proDatosProducto.MetDevolverEnlace(Trim(TxtIdVenta.Text)))
 Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboEstratos()
    On Error GoTo ErrManager
    Set Me.proEstrato = New colEstratoCiudad
    Set Me.proEstrato.proConexion = Me.proConexion
    Me.cboCodigoEstracto.Clear
    Me.cboEstratos.Clear
    If Me.proDatosProducto.proCiudadId <> 0 Then
        If Me.proEstrato.FunGConsulta(Me.proDatosProducto.proCiudadId, 1) Then
            Call SubFPintarComboEstratos
        Else
            Me.cboCodigoEstracto.Clear
            Me.cboEstratos.Clear
            Me.cboCodigoEstracto.ListIndex = -1
            Me.cboEstratos.ListIndex = -1
            MsgBox "Error al consultar los estratos.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarComboEstratos()
    Dim varContador As Integer
    On Error GoTo ErrManager
    Me.cboCodigoEstracto.Clear
    Me.cboEstratos.Clear
    For varContador = 1 To Me.proEstrato.Count
        Me.cboEstratos.AddItem Me.proEstrato.Item(varContador).proNombreEstrato
        Me.cboCodigoEstracto.AddItem Me.proEstrato.Item(varContador).proEstratoCiudadId
    Next varContador
    Me.cboCodigoEstracto.ListIndex = IIf(cboEstratos.ListCount = 1, 0, -1)
    Me.cboEstratos.ListIndex = IIf(cboEstratos.ListCount = 1, 0, -1)
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
 
Private Sub SubFLlenarComboUsos()
    On Error GoTo ErrManager
    Set Me.proUsoServicio = New EDCAdminVoz.colEstratos
    Set Me.proUsoServicio.proConexion = Me.proConexion
    If Me.proUsoServicio.MetConsultar Then
        Call SubFPintarComboUsos
    Else
        Me.cboUso.Clear
        Me.cboCodigoUso.Clear
        Me.cboUso.ListIndex = -1
        Me.cboCodigoUso.ListIndex = -1
        MsgBox "Error al consultar los usos.", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarComboUsos()
    Dim varContador As Integer
    On Error GoTo ErrManager
    Me.cboCodigoUso.Clear
    Me.cboUso.Clear
    For varContador = 1 To Me.proUsoServicio.Count
        Me.cboUso.AddItem Me.proUsoServicio.Item(varContador).proDescripcion
        Me.cboCodigoUso.AddItem Me.proUsoServicio.Item(varContador).proEstratoID
    Next varContador
    Me.cboCodigoUso.ListIndex = IIf(cboUso.ListCount = 1, 0, -1)
    Me.cboUso.ListIndex = IIf(cboUso.ListCount = 1, 0, -1)
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

