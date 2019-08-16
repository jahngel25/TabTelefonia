VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1185
   ClientLeft      =   5910
   ClientTop       =   4875
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSplash 
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      Top             =   1650
      Width           =   5775
      Begin VB.Timer tmrTiempo 
         Interval        =   1000
         Left            =   3420
         Top             =   1650
      End
      Begin VB.Frame fraLogo 
         BackColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   5655
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H004A3869&
            Caption         =   "ADMINISTRACION"
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
            Height          =   225
            Left            =   990
            TabIndex        =   5
            Top             =   330
            Width           =   1995
         End
         Begin VB.Image imgEdificio 
            Height          =   2145
            Left            =   30
            Picture         =   "frmSplash.frx":0000
            Top             =   180
            Width           =   3060
         End
         Begin VB.Image imgATT 
            Height          =   630
            Left            =   3180
            Picture         =   "frmSplash.frx":1561E
            Top             =   120
            Width           =   2400
         End
         Begin VB.Label lblTitulo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sistema de Información de Voz"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   3360
            TabIndex        =   2
            Top             =   930
            Width           =   2115
         End
         Begin VB.Label lblTitulo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sistema de Información de Voz"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   435
            Index           =   1
            Left            =   3390
            TabIndex        =   3
            Top             =   960
            Width           =   2115
         End
         Begin VB.Shape shCuadro 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1665
            Left            =   3210
            Top             =   690
            Width           =   2385
         End
      End
      Begin VB.Shape Shape1 
         Height          =   2715
         Left            =   0
         Top             =   750
         Width           =   5775
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Todos los derechos reservados. CopyRight 2001"
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
         Height          =   405
         Index           =   2
         Left            =   660
         TabIndex        =   4
         Top             =   2580
         Width           =   3915
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1125
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   4545
      _Version        =   65536
      _ExtentX        =   8017
      _ExtentY        =   1984
      _StockProps     =   15
      BackColor       =   12160284
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
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "versión 1.0"
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
         Height          =   225
         Left            =   2820
         TabIndex        =   13
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Administración"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   30
         TabIndex        =   12
         Top             =   870
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D9B548&
         Height          =   225
         Left            =   30
         TabIndex        =   11
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "versión 1.0"
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
         Height          =   225
         Left            =   3150
         TabIndex        =   10
         Top             =   1530
         Width           =   1605
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2340
         Picture         =   "frmSplash.frx":1A520
         Top             =   90
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   30
         Picture         =   "frmSplash.frx":1B1EA
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2190
      End
      Begin VB.Label lblApp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema de Información de Telefonía"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   3705
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D9B548&
         Height          =   225
         Left            =   690
         TabIndex        =   7
         Top             =   630
         Width           =   135
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CRM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D1A316&
         Height          =   765
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         Top             =   90
         Width           =   3105
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo ErrorManager

    Me.lblVersion = "versión " & App.Major & "." & App.Minor & "." & App.Revision
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub



'Propiedad que indica que ventana inicia
Private Sub tmrTiempo_Timer()
On Error GoTo ErrorManager

        tmrTiempo.Enabled = False
        Unload Me
        frmAdminVoz.Show 1
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub
