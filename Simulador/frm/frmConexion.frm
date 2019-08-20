VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmConexion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SONDA de Colombia"
   ClientHeight    =   3345
   ClientLeft      =   2610
   ClientTop       =   4290
   ClientWidth     =   4125
   Icon            =   "frmConexion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4125
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   750
      TabIndex        =   18
      Top             =   120
      Width           =   3315
      Begin VB.Label lblTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Simulador Comunicación EDC - ONYX"
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
         Left            =   120
         TabIndex        =   19
         Top             =   30
         Width           =   3165
      End
   End
   Begin MSForms.Frame fraFX 
      Height          =   315
      Index           =   1
      Left            =   -60
      OleObjectBlob   =   "frmConexion.frx":1CCA
      TabIndex        =   17
      Top             =   90
      Width           =   4785
   End
   Begin MSForms.Frame fraFX 
      Height          =   4125
      Index           =   0
      Left            =   30
      OleObjectBlob   =   "frmConexion.frx":26E2
      TabIndex        =   16
      Top             =   -450
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
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
      Left            =   2910
      TabIndex        =   8
      Top             =   3000
      Width           =   1065
   End
   Begin VB.TextBox Text2 
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
      Left            =   1170
      TabIndex        =   6
      Text            =   "Sonda de Colombia"
      Top             =   2520
      Width           =   2565
   End
   Begin VB.TextBox Text1 
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
      Left            =   1170
      TabIndex        =   4
      Text            =   "11212090"
      Top             =   1800
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Initiate"
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
      Left            =   1830
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
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
      Left            =   1170
      TabIndex        =   5
      Text            =   "0"
      Top             =   2160
      Width           =   825
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1170
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "onyx"
      Top             =   1470
      Width           =   1605
   End
   Begin VB.TextBox txtUser 
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
      Left            =   1170
      TabIndex        =   2
      Text            =   "rcruz"
      Top             =   1110
      Width           =   1575
   End
   Begin VB.TextBox txtServer 
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
      Left            =   1170
      TabIndex        =   0
      Text            =   "COLBTASQL03"
      Top             =   450
      Width           =   2895
   End
   Begin VB.TextBox txtDatabase 
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
      Left            =   1170
      TabIndex        =   1
      Text            =   "OnyxDesarrollo"
      Top             =   780
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   630
      TabIndex        =   15
      Top             =   2580
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "ID Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   420
      TabIndex        =   14
      Top             =   1830
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ID Facturacion"
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
      Top             =   2190
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
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
      TabIndex        =   12
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "User"
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
      Left            =   750
      TabIndex        =   11
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Server"
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
      Left            =   480
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "DataBase"
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
      Left            =   420
      TabIndex        =   9
      Top             =   810
      Width           =   735
   End
End
Attribute VB_Name = "frmConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim varTABVoz As EDCVoz.claONYX
    
    Set varTABVoz = New EDCVoz.claONYX
    
    varTABVoz.DatabaseName = Me.txtDatabase
    varTABVoz.ServerName = Me.txtServer
    varTABVoz.UserLogin = Me.txtUser
    varTABVoz.UserPassword = Me.txtPassword
    varTABVoz.DetailID = Me.txtIDCliente
    varTABVoz.ContactID = Me.Text1
    varTABVoz.ContactName = Me.Text2
    
    varTABVoz.Initiate
End Sub

Private Sub Command2_Click()
    Dim varTABVoz As EDCVoz.claONYX
    
    Set varTABVoz = New EDCVoz.claONYX
    
    varTABVoz.DatabaseName = Me.txtDatabase
    varTABVoz.ServerName = Me.txtServer
    varTABVoz.UserLogin = Me.txtUser
    varTABVoz.UserPassword = Me.txtPassword
    varTABVoz.DetailID = Me.txtIDCliente
    varTABVoz.ContactID = Me.Text1
    varTABVoz.ContactName = Me.Text2
    
    varTABVoz.Load
End Sub

