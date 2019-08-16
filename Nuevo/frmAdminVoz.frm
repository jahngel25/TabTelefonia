VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdminVoz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrador de Información de Voz"
   ClientHeight    =   1185
   ClientLeft      =   2385
   ClientTop       =   990
   ClientWidth     =   16110
   Icon            =   "frmAdminVoz.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   16110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   780
      Width           =   12015
      Begin VB.Timer tmrTiempo 
         Interval        =   1000
         Left            =   2100
         Top             =   0
      End
      Begin MSComctlLib.ImageList imlIconos 
         Left            =   3210
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdminVoz.frx":0CCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdminVoz.frx":19A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdminVoz.frx":267E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdminVoz.frx":3358
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AT&&T Latin America"
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
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblTiempo 
         Alignment       =   2  'Center
         BackColor       =   &H00C09258&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         Top             =   120
         Width           =   1155
      End
   End
   Begin MSComctlLib.Toolbar tlbAdministracion 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16110
      _ExtentX        =   28416
      _ExtentY        =   1376
      ButtonWidth     =   3598
      ButtonHeight    =   1376
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imlIconos"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Parámetros por Producto"
            Object.ToolTipText     =   "Conceptos que restan en una factura"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Valores"
            Object.ToolTipText     =   "Conceptos que suman en una factura"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Seguridad"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Operaciones por novedad"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Administración de Números"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clasificacion"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Productos Relacionados"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reglas"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdminVoz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public proConexion As ADODB.Connection

Private Sub tlbAdministracion_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrorManager

    'Evalua la presión de botón sobre la barra de botones
    Select Case Button.Index
        Case 1 'Parametros por Producto
            Set frmEdicionParametros.proConexion = Me.proConexion
            frmEdicionParametros.Show vbModal
        Case 2 'Valores
            Set frmValor.proConexion = Me.proConexion
            frmValor.Show vbModal
        Case 4 ' Seguridad
            Set frmSeguridad.proConexion = Me.proConexion
            frmSeguridad.Show vbModal
        Case 5 'Operaciones por novedad
            Set frmOperaciones.proConexion = Me.proConexion
            frmOperaciones.Show vbModal
        Case 7 'Administración de Números
            Set frmAdminNumeros.proConexion = Me.proConexion
            frmAdminNumeros.Show (vbModal)
        Case 8 'Clasificacion
            Set frmClasificacion.proConexion = Me.proConexion
            frmClasificacion.Show (vbModal)
        Case 9 'Productos relacionados
            Set frmProductosRelacionados.proConexion = Me.proConexion
            frmProductosRelacionados.Show (vbModal)
         Case 10 'Reglas
            Set FrmRegla.proConexion = Me.proConexion
            FrmRegla.Show (vbModal)
    End Select
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub tmrTiempo_Timer()
    On Error GoTo ErrorManager

        Me.lblTiempo = Format(Now, "Long time")
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub
