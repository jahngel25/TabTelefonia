VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDetailID 
      Height          =   285
      Left            =   1560
      TabIndex        =   18
      Text            =   "189493"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtContactID 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Text            =   "28501"
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtAlternateID 
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
      Left            =   1560
      TabIndex        =   14
      Text            =   "CRM12829851"
      Top             =   0
      Width           =   2895
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
      Left            =   2250
      TabIndex        =   7
      Top             =   5430
      Width           =   1065
   End
   Begin VB.TextBox txtContactName 
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
      Left            =   1560
      TabIndex        =   6
      Text            =   "IMPORTADORA CALI S.A"
      Top             =   750
      Width           =   2925
   End
   Begin VB.TextBox txtUserSite 
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
      Left            =   1560
      TabIndex        =   5
      Text            =   "1"
      Top             =   3000
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
      Left            =   1170
      TabIndex        =   4
      Top             =   5430
      Width           =   1095
   End
   Begin VB.TextBox txtUserPassword 
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
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "onyx"
      Top             =   2640
      Width           =   2925
   End
   Begin VB.TextBox txtUserLogin 
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
      Left            =   1560
      TabIndex        =   2
      Text            =   "euo1848a"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtServerName 
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
      Left            =   1560
      TabIndex        =   1
      Text            =   "colbtadevdbsat"
      Top             =   1920
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
      Left            =   1560
      TabIndex        =   0
      Text            =   "ONYX"
      Top             =   1170
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "DetailID"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label lblContactID 
      Caption         =   "ContactID"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   390
      Width           =   975
   End
   Begin VB.Label lblAlter 
      Caption         =   "AlternateID"
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
      Left            =   360
      TabIndex        =   15
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "ContactName"
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
      Left            =   360
      TabIndex        =   13
      Top             =   750
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "UserSite"
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
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "UserPassword"
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
      Left            =   360
      TabIndex        =   11
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "UserLogin"
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
      Left            =   360
      TabIndex        =   10
      Top             =   2310
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
      Left            =   360
      TabIndex        =   9
      Top             =   1950
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
      Left            =   360
      TabIndex        =   8
      Top             =   1230
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set varTabTel = New EDCVoz.claONYX
    varTabTel.AlternateID = Me.txtAlternateID
    varTabTel.ContactID = Me.txtContactID
    varTabTel.ContactName = Me.txtContactName
    varTabTel.DatabaseName = Me.txtDatabase
    varTabTel.DetailID = Me.txtDetailID
    varTabTel.ServerName = Me.txtServerName
    varTabTel.UserLogin = Me.txtUserLogin
    varTabTel.UserPassword = Me.txtUserPassword
    varTabTel.UserSite = Me.txtUserSite
    
    
    varTabTel.Initiate
End Sub

Private Sub Command2_Click()
Dim varTabTel As EDCVoz.claONYX
    
    Set varTabTel = New EDCVoz.claONYX
    
    varTabTel.AlternateID = Me.txtAlternateID
    varTabTel.ContactID = Me.txtContactID
    varTabTel.ContactName = Me.txtContactName
    varTabTel.DatabaseName = Me.txtDatabase
    varTabTel.DetailID = Me.txtDetailID
    varTabTel.ServerName = Me.txtServerName
    varTabTel.UserLogin = Me.txtUserLogin
    varTabTel.UserPassword = Me.txtUserPassword
    varTabTel.UserSite = Me.txtUserSite
    
    
    varTabTel.Load


End Sub
