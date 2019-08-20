VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProductosRelacionados 
   Caption         =   "Productos Relacionados"
   ClientHeight    =   9195
   ClientLeft      =   480
   ClientTop       =   765
   ClientWidth     =   14550
   Icon            =   "frmProductosRelacionados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14550
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Filtrar por Nombre..."
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
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Modificar Tramo"
      Top             =   420
      Width           =   1755
   End
   Begin VB.CommandButton cmdExcluir 
      BackColor       =   &H00FF8080&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7140
      TabIndex        =   5
      Top             =   3600
      Width           =   345
   End
   Begin VB.CommandButton cmdIncluir 
      BackColor       =   &H00FF8080&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7140
      TabIndex        =   4
      Top             =   3000
      Width           =   345
   End
   Begin VB.Frame FraVariables 
      Caption         =   "Productos Relacionados"
      Height          =   8535
      Index           =   1
      Left            =   7530
      TabIndex        =   2
      Top             =   660
      Width           =   6975
      Begin MSFlexGridLib.MSFlexGrid grdProductos 
         Height          =   8025
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   14155
         _Version        =   393216
         FixedCols       =   0
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
   End
   Begin VB.Frame FraVariables 
      Caption         =   "Productos"
      Height          =   8475
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6975
      Begin VB.Frame fraBusqueda 
         BackColor       =   &H00C09258&
         Caption         =   "  Filtrar por  "
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
         Height          =   1605
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   4605
         Begin VB.CommandButton cmdCerrarBusqueda 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3720
            TabIndex        =   11
            Top             =   0
            Width           =   195
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   210
            MaxLength       =   40
            TabIndex        =   10
            Top             =   540
            Width           =   3495
         End
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "&Mostrar"
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
            Left            =   1380
            TabIndex        =   9
            ToolTipText     =   "Modificar Tramo"
            Top             =   1020
            Width           =   1275
         End
         Begin VB.Label lblNombreDescuento 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Grupo"
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
            Height          =   315
            Left            =   2640
            TabIndex        =   12
            Top             =   330
            Width           =   1395
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdProductosONYX 
         Height          =   8085
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   14261
         _Version        =   393216
         FixedCols       =   0
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEEFE4&
      Caption         =   "Productos Relacionados"
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
      Left            =   7530
      TabIndex        =   7
      Top             =   30
      Width           =   6975
   End
   Begin VB.Label lblNota 
      Alignment       =   2  'Center
      BackColor       =   &H00FEEFE4&
      Caption         =   "Productos de onyx no relacionados"
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
      Left            =   120
      TabIndex        =   6
      Top             =   30
      Width           =   6975
   End
End
Attribute VB_Name = "frmProductosRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' OBJETIVO: Asignar los Productos de Onyx  al tab de telefonia
'****************************************************************
' parametros de entrada: ninguno
' parametros de salida: ninguno
' AUTOR: Hernan Botache
' FECHA: 02/09/2004
'****************************************************************
Option Explicit
'Propiedad de conexion
Public proConexion As ADODB.Connection

Public proProducto As claProductosRelacionados

'Coleccion de productos
Public varProductoOnyx As colProductosRelacionados
Public varProductos As colProductosRelacionados
Dim varFBandera As Integer

Private Sub cmdBuscar_Click()
On Error GoTo ErrorManager

        subFLimpiarBusqueda
        Me.fraBusqueda.Visible = True
        Me.txtNombre.SetFocus
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub
Sub subFLimpiarBusqueda()
On Error GoTo ErrorManager
    Me.txtNombre = ""
    Exit Sub
   
ErrorManager:
    SubGMuestraError
End Sub
Private Sub cmdMostrar_Click()
Dim varError As Boolean
On Error GoTo ErrorManager
Set varProductoOnyx = Nothing
Set varProductoOnyx = New colProductosRelacionados
Set Me.varProductoOnyx.proConexion = Me.proConexion

    If Len(Trim(Me.txtNombre)) Then
        If varProductoOnyx.FunGConsultaxNombre(Trim(Me.txtNombre)) = False Then varError = True
   Else
     If varProductoOnyx.FunGConsultaNoRelacionados = False Then varError = True
   
    End If
      
     subFPintaproductos Me.grdProductosONYX, varProductoOnyx
    'Cierra la ventana de busqueda
    cmdCerrarBusqueda_Click
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
Private Sub cmdCerrarBusqueda_Click()
On Error GoTo ErrorManager

 '   SubGTamObjeto Me.fraBusqueda, 3885, 345, 30
    Me.fraBusqueda.Visible = False
    
    Me.cmdBuscar.SetFocus
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
Private Sub cmdExcluir_Click()
On Error GoTo ErrorManager
    Set proProducto = New claProductosRelacionados
    Set proProducto.proConexion = Me.proConexion
    Set Me.proProducto.proColProductosRelacionados = Nothing
    Set Me.proProducto.proColProductosRelacionados = varProductos

    If Me.grdProductos.Row = 0 Then
        MsgBox "Debe seleccionar un producto a excluir", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
   ' adiciona el item a los productos sin asignar
   varProductoOnyx.Add varProductos.Item(Me.grdProductos.Row).proConexion, varProductos.Item(Me.grdProductos.Row).proProductNumber, varProductos.Item(Me.grdProductos.Row).provchDescription
   'elimina el item de los productos asignados
   If Me.proProducto.FunGEliminarProducto(varProductos.Item(Me.grdProductos.Row)) = False Then
            MsgBox "No fue posible eliminar el Producto", vbInformation + vbOKOnly, App.Title
            Exit Sub
    End If
    Set varProductos = Me.proProducto.proColProductosRelacionados
    Set proProducto = Nothing
    'Muestra los productos asignados
    subFPintaproductos Me.grdProductos, varProductos
   
    'Pinta todos los productos sin asignar
    subFPintaproductos Me.grdProductosONYX, varProductoOnyx
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdIncluir_Click()
On Error GoTo ErrorManager
    Set proProducto = New claProductosRelacionados
    Set proProducto.proConexion = Me.proConexion
    Set Me.proProducto.proColProductosRelacionados = Nothing
    Set Me.proProducto.proColProductosRelacionados = varProductos
'
    If Me.grdProductosONYX.Row = 0 Then
        MsgBox "Debe seleccionar un producto a ser incluido de los productos relacionados", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    
    'Agrega el grupo a los productos asignados
    If Me.proProducto.FunGAgregarProducto(varProductoOnyx.Item(Me.grdProductosONYX.Row)) = False Then
            MsgBox "No fue posible agregar el producto", vbInformation + vbOKOnly, App.Title
            Exit Sub
    End If
    Set varProductos = Me.proProducto.proColProductosRelacionados
    'remueve el grupo de los productos sin asignar
    varProductoOnyx.Remove (Me.grdProductosONYX.Row)
    subFPintaproductos Me.grdProductos, varProductos
   
    subFPintaproductos Me.grdProductosONYX, varProductoOnyx
    Set proProducto = Nothing
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorManager

    If varFBandera Then
        varFBandera = 0
        Unload Me
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
Sub subFPintaproductos(parGrid As MSFlexGrid, parColProductosrelacionados As colProductosRelacionados, Optional parCadena As Variant)

Dim varCuenta As Integer
Dim varProductos As claProductosRelacionados
Dim varTamaño As Integer
On Error GoTo ErrorManager
    
    parGrid.Redraw = False
    
    'Adecuación de la grilla
    parGrid.Rows = 2
    parGrid.Cols = 2
    parGrid.TextMatrix(0, 0) = "Producto"
    parGrid.FixedRows = 1
    parGrid.ColWidth(0) = 5700
    parGrid.ColWidth(1) = 0
    
    parGrid.Rows = 1
    
    'Busca el tamaño de la cadena
    If IsMissing(parCadena) = False Then
        varTamaño = Len(Trim(parCadena))
    End If
    
    'Recorre la coleccion para llenar la grilla
    For varCuenta = 1 To parColProductosrelacionados.Count
    
            'Instancia del objeto
            Set varProductos = Nothing
            Set varProductos = New claProductosRelacionados
            
            'Copia los datos a la clase aislada
            varProductos.provchDescription = parColProductosrelacionados.Item(varCuenta).provchDescription
            
            'Agrega a la grilla
            parGrid.AddItem varProductos.provchDescription
                            
            
            If IsMissing(parCadena) = False Then
                    'Si no cumple el criterio de cadena, oculta la fila
                    If UCase(Trim(parCadena)) <> UCase(Left(varProductos.provchDescription, varTamaño)) Then
                            parGrid.RowHeight(parGrid.Rows - 1) = 0
                    End If
            End If
    Next varCuenta
    
    'Desinstancia del producto
    Set varProductos = Nothing
    
     parGrid.Redraw = True
    parGrid.Refresh
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()

On Error GoTo ErrorManager
        Set varProductoOnyx = New colProductosRelacionados
        Set varProductoOnyx.proConexion = Me.proConexion
        Set varProductos = New colProductosRelacionados
        Set varProductos.proConexion = Me.proConexion
        
       'Consulta de los productos de onyx que no han sido asignados
    If varProductoOnyx.FunGConsultaNoRelacionados = False Then
        MsgBox "No fue posible realizar la consulta de variables de ONYX", vbInformation + vbOKOnly, App.Title
        varFBandera = 1
        Exit Sub
    End If
    'consulta los productos que han sido asignados
    If varProductos.FunGConsulta = False Then
        MsgBox "No fue posible realizar la consulta de variables de ONYX", vbInformation + vbOKOnly, App.Title
        varFBandera = 1
        Exit Sub
    End If
    'Despliega los productos en las grillas
    subFPintaproductos Me.grdProductosONYX, varProductoOnyx
    subFPintaproductos Me.grdProductos, varProductos
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorManager

    varFBandera = 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub grdProductos_DblClick()
On Error GoTo ErrorManager

    cmdExcluir_Click
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub grdProductosOnyx_DblClick()
On Error GoTo ErrorManager

    cmdIncluir_Click
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeAlfaNumerico(KeyAscii, 0)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
