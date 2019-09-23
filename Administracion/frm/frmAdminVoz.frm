VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdminVoz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Administración del Sistema de Telefonía"
   ClientHeight    =   2145
   ClientLeft      =   315
   ClientTop       =   1995
   ClientWidth     =   13635
   Icon            =   "frmAdminVoz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   13635
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   2
      Top             =   1620
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   926
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21431
            Text            =   "TELMEX Colombia"
            TextSave        =   "TELMEX Colombia"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "2:33 p. m."
         EndProperty
      EndProperty
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
   Begin MSComctlLib.Toolbar tlbAdministracion1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   1482
      ButtonWidth     =   3757
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlIconos"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Inserción de Números"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consulta de Números"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clasificación de Números"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Usuario Clasificación"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reglas de clasificación"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Serv. Suplementarios"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Valores Serv. Suple"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.Toolbar tlbAdministracion 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   1482
      ButtonWidth     =   3704
      ButtonHeight    =   1429
      Wrappable       =   0   'False
      Appearance      =   1
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
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prod. Relacionados"
            Object.ToolTipText     =   "Conceptos que suman en una factura"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Operaciones x Novedad"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Seguridad"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Estratos x Ciudad"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Normas"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imlIconos 
      Left            =   0
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
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
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminVoz.frx":4032
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminVoz.frx":4D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminVoz.frx":59E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminVoz.frx":66C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminVoz.frx":739A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminVoz.frx":8074
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminVoz.frx":8D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminVoz.frx":99A0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdminVoz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Se agrega el seteo de proLogin cuando se
'       invoca el form consulta de numeros
'       REQUERIMIENTO:          5322
'       VERSION:       1.0.000
'       FECHA:       2009/10/10
'*******************************************************************
Option Explicit
Public varColUsuarios As colUsuario
Public proConexion As ADODB.Connection
Public proLogin As String


Sub SubFSeguridad()
Dim varCuenta As Integer
Dim varEncontro As Boolean
On Error GoTo ErrorManager

    'Por default oculta el botón de administración
    Me.tlbAdministracion.Visible = False
    Me.tlbAdministracion1.Visible = False
    
    'instancia la colección de usuarios
    Set varColUsuarios = New colUsuario
    Set varColUsuarios.proConexion = Me.proConexion
    varColUsuarios.proAplicacionId = AplicacionID
    
    Me.Height = 2550 - 810
    'Revisa los usuarios autorizados
    If varColUsuarios.FunGConsultaxApp Then
        varCuenta = 1
        varEncontro = False
        varGAdminAplicacion = False
        varGAdminTelefonia = False
        While varCuenta <= varColUsuarios.Count And varEncontro = False
            If Trim(varColUsuarios(varCuenta).proUserId) = Trim(Me.proLogin) Then
                'Busca los permisos
                varColUsuarios(varCuenta).proPrivilegios = Trim(varColUsuarios(varCuenta).proPrivilegios)
                varGAdminAplicacion = (Left(varColUsuarios(varCuenta).proPrivilegios, 1) = "1")
                varGAdminTelefonia = (Mid(varColUsuarios(varCuenta).proPrivilegios, 2, 1) = "1")
                varEncontro = True
            Else
                varCuenta = varCuenta + 1
            End If
        Wend
        
        'Permisos para ver el botón de administración
        If varGAdminTelefonia Then
            Me.tlbAdministracion1.Visible = True
        End If
        
        If varGAdminAplicacion Then
            Me.tlbAdministracion.Visible = True
        End If
        
        If varGAdminAplicacion And varGAdminTelefonia Then
                Me.Height = 2550
        End If
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


Private Sub Form_Load()
On Error GoTo ErrorManager

    Call SubFSeguridad
    Exit Sub
        
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub tlbAdministracion_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrorManager

    'Evalua la presión de botón sobre la barra de botones
    Select Case Button.index
        Case 1 'Parametros por Producto
            Set frmEdicionParametros.proConexion = Me.proConexion
            frmEdicionParametros.Show vbModal
        Case 2 'Valores
            Set frmValor.proConexion = Me.proConexion
            frmValor.proPermitirInsertar = False
            frmValor.Show vbModal
        Case 4 'Productos Relacionados
            Set frmProductosRelacionados.proConexion = Me.proConexion
            frmProductosRelacionados.Show (vbModal)
        Case 5 'Operaciones por novedad
            Set frmOperaciones.proConexion = Me.proConexion
            frmOperaciones.Show vbModal
        Case 7 ' Seguridad
            Set frmSeguridad.proConexion = Me.proConexion
            frmSeguridad.Show vbModal
        Case 9 ' Estratos
            Set frmEstratoCiudad.proConexion = Me.proConexion
            frmEstratoCiudad.Show vbModal
        Case 10 ' Normas
            Set frmNorma.proConexion = Me.proConexion
            frmNorma.Show vbModal
    End Select
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub tlbAdministracion1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrorManager

    'Evalua la presión de botón sobre la barra de botones
    Select Case Button.index
        Case 1 'Administración de Números
            Set frmAdminNumeros.proConexion = Me.proConexion
            frmAdminNumeros.proUsuario = Me.proLogin
            frmAdminNumeros.Show (vbModal)
        Case 2 'Consulta de Números
            Set frmConsultaNumeros.proConexion = Me.proConexion
            '/* 1.0.000  -  Inicio */
            frmConsultaNumeros.proLogin = Me.proLogin
            '/* 1.0.000  -  Fin */
            frmConsultaNumeros.proLlamadoAdministracion = True
            frmConsultaNumeros.Show vbModal
            frmConsultaNumeros.proLlamadoAdministracion = False
        Case 4 'Clasificacion
            Set frmClasificacion.proConexion = Me.proConexion
            frmClasificacion.Show (vbModal)
        Case 5 'Usuarios Clasificacion
            Set frmUsersClasificacion.proConexion = Me.proConexion
            frmUsersClasificacion.Show (vbModal)
        Case 6 'Reglas
            Set FrmRegla.proConexion = Me.proConexion
            FrmRegla.Show (vbModal)
         Case 8 'Serv. Suplementarios
            Set frmServiciosSuplementarios.proConexion = Me.proConexion
            frmServiciosSuplementarios.Show vbModal
         Case 9 'Valores Serv. Suplementarios
            Set frmValoresServSuplementarios.proConexion = Me.proConexion
            frmValoresServSuplementarios.Show vbModal
    End Select
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
