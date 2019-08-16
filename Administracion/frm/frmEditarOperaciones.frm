VERSION 5.00
Begin VB.Form frmEditarOperaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edición de Operaciones por novedad"
   ClientHeight    =   2985
   ClientLeft      =   4770
   ClientTop       =   4560
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   5370
      Begin VB.ComboBox CmbSeccionNombre 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "Debe seleccionar un tipo de sección"
         Top             =   1680
         Width           =   4200
      End
      Begin VB.ComboBox CmbSeccionCodigo 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cmbCategoriaNombre 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Debe seleccionar una categoria"
         Top             =   360
         Width           =   4200
      End
      Begin VB.ComboBox cmbTipoNombre 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Debe seleccionar un tipo"
         Top             =   765
         Width           =   4200
      End
      Begin VB.ComboBox cmbTiposNovedadNombre 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Debe seleccionar una operación"
         Top             =   1215
         Width           =   4200
      End
      Begin VB.ComboBox cmbTiposNovedadCodigo 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1215
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cmbTipoCodigo 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   765
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cmbCategoriaCodigo 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Sección"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria"
         Height          =   240
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo"
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   810
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Operación"
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   1215
         Width           =   870
      End
   End
   Begin VB.CommandButton btnGuardar 
      Caption         =   "&Guardar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   1425
   End
End
Attribute VB_Name = "frmEditarOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public proOperacionOnyx As claOperacionOnyx
Public proTiposNovedad As colTiposNovedad
Public proConexion As ADODB.Connection
Dim varNuevo As Boolean
Sub LlenarComboCategoria()
    Me.cmbCategoriaNombre.AddItem "OT's"
    Me.cmbCategoriaNombre.AddItem "Atenciones"
    Me.cmbCategoriaNombre.AddItem "Ventas"
    Me.cmbCategoriaCodigo.AddItem "1"
    Me.cmbCategoriaCodigo.AddItem "2"
    Me.cmbCategoriaCodigo.AddItem "3"
End Sub
Private Sub SubFPintarComboTipos()
    Dim varContador As Integer
    Dim varIncidentType As colIncidentType
    On Error GoTo ErrManager
    
    If Me.cmbCategoriaCodigo.Text <> "" Then
        Me.cmbTipoCodigo.Clear
        Me.cmbTipoNombre.Clear
        Set varIncidentType = New colIncidentType
        Set varIncidentType.proConexion = Me.proConexion
        varIncidentType.proParentId = Me.cmbCategoriaCodigo.Text
        varIncidentType.MetConsultar
        For varContador = 1 To varIncidentType.Count
            Me.cmbTipoNombre.AddItem varIncidentType.Item(varContador).provchParameterDesc
            Me.cmbTipoCodigo.AddItem varIncidentType.Item(varContador).proParameterId
        Next varContador
        
        Me.cmbTipoCodigo.ListIndex = -1
        Me.cmbTipoNombre.ListIndex = -1
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Sub LlenarComboOperacion()
'De la Tabla TiposNovedad
    On Error GoTo ErrManager
    
    Set Me.proTiposNovedad = New colTiposNovedad
    Set Me.proTiposNovedad.proConexion = Me.proConexion
    
    If Me.proTiposNovedad.MetConsultar Then
        Call SubFPintarComboOperacion
    Else
        MsgBox "Error al consultar los TiposNovedads.", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarComboOperacion()
    Dim varContador As Integer
    Dim proTiposCategoria As colTiposNovedad
    On Error GoTo ErrManager
    Me.cmbTiposNovedadCodigo.Clear
    Me.cmbTiposNovedadNombre.Clear
    Set proTiposCategoria = New colTiposNovedad
    Set proTiposCategoria.proConexion = Me.proConexion
    proTiposCategoria.proTipoNovedadId = Me.cmbCategoriaCodigo.Text
    proTiposCategoria.MetConsultar
    For varContador = 1 To proTiposCategoria.Count
        Me.cmbTiposNovedadNombre.AddItem proTiposCategoria.Item(varContador).proDescripcionNovedad
        Me.cmbTiposNovedadCodigo.AddItem proTiposCategoria.Item(varContador).proTipoNovedadId
    Next varContador
    
    Me.cmbTiposNovedadCodigo.ListIndex = -1
    Me.cmbTiposNovedadNombre.ListIndex = -1
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Function FunFCopiaDatosaClase() As Boolean
Dim varCadena As String
On Error GoTo ErrorManager

        Me.proOperacionOnyx.proIncidentCategory = Me.cmbCategoriaCodigo.Text
        Me.proOperacionOnyx.proIncidentTypeId = Me.cmbTipoCodigo.Text
        Me.proOperacionOnyx.proTipoNovedadId = Me.cmbTiposNovedadCodigo.Text
        Me.proOperacionOnyx.proTipoSeccionId = Me.CmbSeccionCodigo
        
        FunFCopiaDatosaClase = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
Private Sub cmdGuardar_Click()
On Error GoTo ErrorManager
        
        'Valida que exista un usuarios
        If Me.cmbCategoriaCodigo.ListIndex = -1 Or Me.cmbTipoCodigo.ListIndex = -1 Or Me.cmbTiposNovedadCodigo.ListIndex = -1 Then
                MsgBox "Es indispensable indicar un usuario válido. Los usuarios que ya pertenecen a la aplicación no podrán ser reingresados.", vbInformation, App.Title
                Exit Sub
        End If

        FunFCopiaDatosaClase
        
            If Me.proOperacionOnyx.MetInsertar = False Then
                    MsgBox "No fue posible agregar al usuario", vbInformation, App.Title
            End If
        
        MsgBox "Se almacenó exitosamente. Los cambios tendrán efecto la siguiente vez que el usuario ingrese a la aplicación", vbInformation, App.Title
        
        'Descarga la forma
        Unload Me
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub cmbCategoria_Change()

End Sub

Private Sub btnGuardar_Click()
On Error GoTo ErrorManager
    
    If FunGVerificaDatos(Me.cmbCategoriaNombre) = False Then Exit Sub
    If FunGVerificaDatos(Me.cmbTipoNombre) = False Then Exit Sub
    If FunGVerificaDatos(Me.cmbTiposNovedadNombre) = False Then Exit Sub
    If FunGVerificaDatos(Me.CmbSeccionNombre) = False Then Exit Sub

    'Copia los datos de controles a las propiedades de la clase
    Call FunFCopiaDatosaClase
    
    'Almacena en la base
    If Me.proOperacionOnyx.MetInsertar = False Then
        MsgBox "No fue posible guardar los cambios", vbInformation + vbOKOnly, App.Title
    End If
    
   'Unload Me
Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmbCategoriaNombre_Click()
        Me.cmbCategoriaCodigo.ListIndex = Me.cmbCategoriaNombre.ListIndex
        Me.cmbCategoriaCodigo.Text = cmbCategoriaCodigo.List(Me.cmbCategoriaNombre.ListIndex)
        Call SubFPintarComboTipos
End Sub

Private Sub cmbTipoNombre_Click()
    Me.cmbTipoCodigo.ListIndex = Me.cmbTipoNombre.ListIndex
End Sub

Private Sub cmbTiposNovedadNombre_Click()
    Me.cmbTiposNovedadCodigo.ListIndex = Me.cmbTiposNovedadNombre.ListIndex
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager

        'Consulta la colección de usuarios
        Call LlenarComboCategoria
        Call LlenarComboOperacion
        Call LlenarComboSeccion
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub
Sub LlenarComboSeccion()
    Me.CmbSeccionNombre.AddItem "Tipos de Linea"
    Me.CmbSeccionNombre.AddItem "Numeración Publica"
    Me.CmbSeccionNombre.AddItem "Numeración Coorporativa"
    Me.CmbSeccionNombre.AddItem "Todos"
    Me.CmbSeccionCodigo.AddItem "T"
    Me.CmbSeccionCodigo.AddItem "P"
    Me.CmbSeccionCodigo.AddItem "C"
    Me.CmbSeccionCodigo.AddItem "*"
End Sub
Private Sub cmbSeccionNombre_Click()
    Me.CmbSeccionCodigo.ListIndex = Me.CmbSeccionNombre.ListIndex
End Sub



