VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Administracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion"
   ClientHeight    =   2265
   ClientLeft      =   2820
   ClientTop       =   1875
   ClientWidth     =   4260
   Icon            =   "Administracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4260
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
      Left            =   1230
      TabIndex        =   9
      Text            =   "onyxdesarrollo"
      Top             =   750
      Width           =   2895
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
      Left            =   1230
      TabIndex        =   8
      Text            =   "ATTLAMSD-02"
      Top             =   360
      Width           =   2895
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
      Left            =   1230
      TabIndex        =   7
      Text            =   "sa"
      Top             =   1080
      Width           =   1575
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
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   6
      Text            =   "onyx"
      Top             =   1440
      Width           =   1605
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
      Left            =   1230
      TabIndex        =   5
      Top             =   2130
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton Command1 
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
      Height          =   345
      Left            =   1980
      TabIndex        =   4
      Top             =   1830
      Width           =   2205
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
      Left            =   1230
      TabIndex        =   3
      Text            =   "34906"
      Top             =   1770
      Visible         =   0   'False
      Width           =   825
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
      Left            =   1230
      TabIndex        =   2
      Text            =   "TELECOM"
      Top             =   2490
      Visible         =   0   'False
      Width           =   2565
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
      Left            =   2970
      TabIndex        =   1
      Top             =   2970
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSForms.Frame fraFX 
      Height          =   4125
      Index           =   0
      Left            =   90
      OleObjectBlob   =   "Administracion.frx":0CCA
      TabIndex        =   0
      Top             =   -480
      Width           =   315
   End
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
      Left            =   900
      TabIndex        =   17
      Top             =   60
      Width           =   3165
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
      Left            =   480
      TabIndex        =   16
      Top             =   780
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
      Left            =   540
      TabIndex        =   15
      Top             =   450
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
      Left            =   810
      TabIndex        =   14
      Top             =   1110
      Width           =   735
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
      Left            =   480
      TabIndex        =   13
      Top             =   1470
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
      Left            =   480
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   480
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
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
      Left            =   690
      TabIndex        =   10
      Top             =   2550
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "Administracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim varFacturacion As EDCAdminFacturacion.claONYX
    
    Set varFacturacion = New EDCAdminFacturacion.claONYX
    
    varFacturacion.DatabaseName = Me.txtDatabase
    varFacturacion.ServerName = Me.txtServer
    varFacturacion.UserLogin = Me.txtUser
    varFacturacion.UserPassword = Me.txtPassword
    varFacturacion.DetailID = Me.txtIDCliente
    varFacturacion.ContactID = Me.Text1
    varFacturacion.ContactName = Me.Text2
    
    varFacturacion.Initiate
End Sub

Private Sub Command2_Click()
    Dim varFacturacion As EDCAdminFacturacion.claONYX
    
    Set varFacturacion = New EDCAdminFacturacion.claONYX
    
    varFacturacion.DatabaseName = Me.txtDatabase
    varFacturacion.ServerName = Me.txtServer
    varFacturacion.UserLogin = Me.txtUser
    varFacturacion.UserPassword = Me.txtPassword
    varFacturacion.DetailID = Me.txtIDCliente
    varFacturacion.ContactID = Me.Text1
    varFacturacion.ContactName = Me.Text2
    
    varFacturacion.Load
End Sub

