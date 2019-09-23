VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   1185
   ClientLeft      =   4725
   ClientTop       =   4680
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   1125
      Left            =   30
      TabIndex        =   2
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
      Begin VB.Label Label2 
         BackColor       =   &H00D9B548&
         Height          =   225
         Left            =   720
         TabIndex        =   7
         Top             =   630
         Width           =   135
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
         Left            =   840
         TabIndex        =   5
         Top             =   600
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   30
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2190
      End
      Begin VB.Image imgEdificio 
         Height          =   480
         Left            =   2340
         Picture         =   "frmSplash.frx":13F2
         Top             =   90
         Width           =   480
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
         Index           =   0
         Left            =   1890
         TabIndex        =   4
         Top             =   90
         Width           =   3105
      End
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
         Left            =   2490
         TabIndex        =   3
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D9B548&
         Height          =   225
         Left            =   30
         TabIndex        =   6
         Top             =   630
         Width           =   645
      End
   End
   Begin VB.Frame fraSplash 
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   5775
      Begin Threed.SSPanel SSPanel2 
         Height          =   1365
         Left            =   0
         TabIndex        =   8
         Top             =   90
         Width           =   4785
         _Version        =   65536
         _ExtentX        =   8440
         _ExtentY        =   2408
         _StockProps     =   15
         Caption         =   "SSPanel2"
         BackColor       =   0
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
      End
      Begin VB.Timer tmrTiempo 
         Interval        =   2000
         Left            =   2700
         Top             =   1620
      End
      Begin VB.Frame fraLogo 
         BackColor       =   &H00B98D1C&
         Height          =   1305
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   4335
         Begin VB.Shape Shape1 
            BackColor       =   &H00B98D1C&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1185
            Left            =   270
            Top             =   1410
            Width           =   2385
         End
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00808080&
         Height          =   1365
         Left            =   0
         Top             =   90
         Width           =   4395
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

Private Sub tmrTiempo_Timer()
On Error GoTo ErrorManager

        tmrTiempo.Enabled = False
        Unload Me
        frmVoz.Show 1
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub
