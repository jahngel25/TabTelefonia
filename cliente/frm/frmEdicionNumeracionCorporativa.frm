VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEdicionNumeracionCorporativa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de la Numeración Corporativa"
   ClientHeight    =   5295
   ClientLeft      =   5385
   ClientTop       =   4935
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5190
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CheckBox chkVirtual 
      Caption         =   "Virtual"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtNumero 
      Height          =   345
      Left            =   2760
      TabIndex        =   0
      Top             =   60
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid grdNumeracionCorporativa 
      Height          =   3585
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   6324
      _Version        =   393216
   End
   Begin VB.Frame fraTituloProducto 
      BackColor       =   &H00C09258&
      Caption         =   "Numeros a insertar "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5145
   End
   Begin VB.Label lblNumero 
      AutoSize        =   -1  'True
      Caption         =   "Número a insertar:"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   150
      Width           =   1290
   End
End
Attribute VB_Name = "frmEdicionNumeracionCorporativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proDatosProducto As claDatosProducto
Public proOnyx As EDCVoz.claONYX

Private varNovedadNumeracionCorporativa As claNovedadNumeracionCorporativa

Public proConexion As ADODB.Connection

Private Sub cmdGuardar_Click()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    If Trim(Me.txtNumero.Text) = "" Then
        MsgBox "Debe digitar el número a ingresar.", vbInformation, App.Title
        Exit Sub
    End If
    
    For varContador = 1 To Me.proDatosProducto.proNovedadNumeracionCorporativa.Count
        If Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proIncidentId = _
           Me.proDatosProducto.proIncidentId And _
           Trim(Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proMarcacion) = _
           Trim(Me.txtNumero.Text) Then
            MsgBox "El número digitado ya existe para el incidente actual.", vbInformation, App.Title
            Exit Sub
        End If
    Next varContador
    
    Set varNovedadNumeracionCorporativa = Nothing
    Set varNovedadNumeracionCorporativa = New claNovedadNumeracionCorporativa
    Set varNovedadNumeracionCorporativa.proConexion = Me.proConexion
    
    'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
    If Not Me.proDatosProducto.MetGuardar Then
        MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Inserta o actualiza la información de los incidentes
    If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
        MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
    End If
    
    varNovedadNumeracionCorporativa.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
    varNovedadNumeracionCorporativa.proIncidentId = Me.proDatosProducto.proIncidentId
    varNovedadNumeracionCorporativa.proTipoNovedadId = "1"
    varNovedadNumeracionCorporativa.proMarcacion = Trim(Me.txtNumero.Text)
    varNovedadNumeracionCorporativa.proVirtual = IIf(Me.chkVirtual.Value = 1, "S", "N")
    
    If varNovedadNumeracionCorporativa.FunGInsertar Then
        If Me.proDatosProducto.MetAgregarNovedadNumeracionCorporativa(varNovedadNumeracionCorporativa) Then
            Call SubFPintarGridNumeros
            Me.txtNumero.Text = ""
            Me.chkVirtual.Value = 0
            Me.txtNumero.SetFocus
        Else
            MsgBox "Error al agregar el número.", vbCritical, App.Title
            Exit Sub
        End If
    Else
        MsgBox "Error al ingresar el número.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    
    'Inicializar el grid
    Call SubFInicializarGridNumeros
    
    'Pintar el grid
    Call SubFPintarGridNumeros
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridNumeros()
    On Error GoTo ErrManager
    
    With Me.grdNumeracionCorporativa
        .Rows = 1
        .Cols = 5
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
        .TextMatrix(0, 2) = "TipoNovedad"
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = "Marcación"
        
        'Columna agregada por Carlos Castelblanco 2006/07/26
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Virtual"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridNumeros()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdNumeracionCorporativa.Rows = 1
    
    For varContador = 1 To Me.proDatosProducto.proNovedadNumeracionCorporativa.Count
        Me.grdNumeracionCorporativa.AddItem Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proDatosProductoId & vbTab & _
                                            Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proIncidentId & vbTab & _
                                            Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proTipoNovedadId & vbTab & _
                                            Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proMarcacion & vbTab & _
                                            Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proVirtual
                                            'Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proVirtual  Agregado por CArlos Castelblanco 2006/07/26
        
        If Me.proDatosProducto.proNovedadNumeracionCorporativa.Item(varContador).proTipoNovedadId <> "1" Then
            Me.grdNumeracionCorporativa.RowHeight(Me.grdNumeracionCorporativa.Rows - 1) = 0
        End If
                                            
    Next varContador
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
