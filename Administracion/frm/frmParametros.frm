VERSION 5.00
Begin VB.Form frmEdicionParametros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración de Parámetros"
   ClientHeight    =   7755
   ClientLeft      =   2250
   ClientTop       =   2235
   ClientWidth     =   10875
   Icon            =   "frmParametros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerar 
      Height          =   315
      Left            =   5820
      Picture         =   "frmParametros.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame fraFondoFiltro 
      Height          =   555
      Left            =   60
      TabIndex        =   7
      Top             =   180
      Width           =   10695
      Begin VB.ComboBox cboCodigoProducto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   150
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cboNombreProducto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   9345
      End
      Begin VB.Label lblProducto 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
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
         Left            =   300
         TabIndex        =   9
         Top             =   210
         Width           =   690
      End
   End
   Begin EDCAdminVoz.ctlLstJerarquia lstDefinicion 
      Height          =   3495
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   780
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   6165
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MostrarBotonesV =   -1  'True
      MostrarBotonesH =   0   'False
      Enabled         =   0   'False
   End
   Begin EDCAdminVoz.ctlLstJerarquia lstDefinicion 
      Height          =   3495
      Index           =   1
      Left            =   3645
      TabIndex        =   2
      Top             =   780
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   6165
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MostrarBotonesV =   -1  'True
      MostrarBotonesH =   0   'False
      Enabled         =   0   'False
   End
   Begin EDCAdminVoz.ctlLstJerarquia lstDefinicion 
      Height          =   3495
      Index           =   2
      Left            =   7200
      TabIndex        =   3
      Top             =   780
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   6165
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MostrarBotonesV =   -1  'True
      MostrarBotonesH =   0   'False
      Enabled         =   0   'False
   End
   Begin EDCAdminVoz.ctlLstJerarquia lstDatos 
      Height          =   3315
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   4260
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   5847
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MostrarBotonesV =   -1  'True
      MostrarBotonesH =   0   'False
      Enabled         =   0   'False
   End
   Begin EDCAdminVoz.ctlLstJerarquia lstDatos 
      Height          =   3315
      Index           =   1
      Left            =   3630
      TabIndex        =   5
      Top             =   4260
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   5847
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MostrarBotonesV =   -1  'True
      MostrarBotonesH =   0   'False
      Enabled         =   0   'False
   End
   Begin EDCAdminVoz.ctlLstJerarquia lstDatos 
      Height          =   3315
      Index           =   2
      Left            =   7170
      TabIndex        =   6
      Top             =   4260
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   5847
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MostrarBotonesV =   -1  'True
      MostrarBotonesH =   0   'False
      Enabled         =   0   'False
   End
   Begin VB.Frame FraFiltro 
      BackColor       =   &H00C09258&
      Caption         =   "  Filtro "
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
      Height          =   7725
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10845
   End
End
Attribute VB_Name = "frmEdicionParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmEdicionParametros
' Fecha  : 23/09/2004 17:45
' Autor    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : Forma que permite la configuración de parámetros para los detalles de productos de Telefonía
'---------------------------------------------------------------------------------------

Option Explicit

'Propiedad de conexion
Public proConexion As ADODB.Connection
Public proProductNumber As String

Public proParametroProducto As claParametroProducto

'Colecciones
Public procolParametroProducto As colParametroProducto
Public proclaParametroProducto As claParametroProducto

Public procolValordatos As colValordatos
Public proclaValordatos As claValordatos

Public procolValoresCampoProducto As colValoresCampoProducto
Public proclaValoresCampoProducto As claValoresCampoProducto

'Header
Public proParametrosProducto As colParametroProducto
Public proProducto As colProductMaster
Public proValorId As Long
'Constantes
Dim bIniciada As Boolean
Const TipoTelefoniaLocal = "1810"
Const TipoTelefoniaNacional = "1811"
Const sCampoSubred = "vchUser7"
Const sCampoDireccionIP = "vchUser4"

Public proCampoPadre1 As String
Public proCampoPadre2 As String
Public proCampoPadre3 As String
Public proValorPadre1 As String
Public proValorPadre2 As String

Public proEdicionCliente As Boolean
Private Sub AjustarvaloresEdicionCliente()
   On Error GoTo ErrorManager
    If Not proEdicionCliente Then Exit Sub
        
    If proCampoPadre1 <> "" Then
        Call BuscarListaDefinicion(0, proCampoPadre1)
    End If
    
    If proCampoPadre2 <> "" Then
        Call BuscarListaDefinicion(1, proCampoPadre2)
    End If
    
    If proValorPadre1 <> "" Then
            Call BuscarListaValor(0, proValorPadre1)
    End If
    
    If proCampoPadre3 <> "" Then
        Call BuscarListaDefinicion(2, proCampoPadre3)
    End If
    
    If proValorPadre2 <> "" Then
            Call BuscarListaValor(1, proValorPadre2)
    End If
    
    Exit Sub
      
ErrorManager:
    SubGMuestraError
    
End Sub
Sub BuscarListaDefinicion(Index As Integer, sValor As String)

   On Error GoTo ErrorManager

        Dim i As Integer
        For i = 0 To lstDefinicion(Index).ListCount
            If lstDefinicion(Index).ListCodigo(i) = sValor Then
                lstDefinicion(Index).ListIndex = i
                Exit For
            End If
        Next
        lstDefinicion(Index).Enabled = False
        
      Exit Sub

ErrorManager:
    SubGMuestraError
End Sub
Sub BuscarListaValor(Index As Integer, sValor As String)

   On Error GoTo ErrorManager

        Dim i As Integer
        For i = 0 To lstDatos(Index).ListCount
            If lstDatos(Index).ListCodigo(i) = sValor Then
                lstDatos(Index).ListIndex = i
                Exit For
            End If
        Next
        lstDatos(Index).Enabled = False
      Exit Sub

ErrorManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboProductos()
    
    On Error GoTo ErrorManager
    
    Set Me.proProducto = New colProductMaster
    Set Me.proProducto.proConexion = Me.proConexion
    
    If Me.proProducto.MetConsultar Then
        Call SubFPintarComboProductos
    Else
        MsgBox "Error al consultar los productos.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarComboProductos()
    
    Dim varContador As Integer
    On Error GoTo ErrorManager
    
    Me.cboCodigoProducto.Clear
    Me.cboNombreProducto.Clear
    
    For varContador = 1 To Me.proProducto.Count
        Me.cboNombreProducto.AddItem Me.proProducto.Item(varContador).proDescription
        Me.cboCodigoProducto.AddItem Me.proProducto.Item(varContador).proProductNumber
    Next varContador
    
    If proParametroProducto Is Nothing Then
        Me.cboCodigoProducto.ListIndex = -1
        Me.cboNombreProducto.ListIndex = -1
    Else
        If proParametroProducto.proProductNumber <> "" Then
            Dim i As Integer
            For i = 0 To cboCodigoProducto.ListCount
                If cboCodigoProducto.List(i) = Trim(proParametroProducto.proProductNumber) Then
                    cboCodigoProducto.ListIndex = i
                    Me.cboNombreProducto.ListIndex = i
                    Exit For
                End If
            Next
        End If
    End If
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Sub IniciarClases()
    On Error GoTo ErrorManager
    Set proclaParametroProducto = New claParametroProducto
    Set proclaParametroProducto.proConexion = Me.proConexion
        
    Set proclaValordatos = New claValordatos
    Set proclaValordatos.proConexion = Me.proConexion
    
    Set proclaValoresCampoProducto = New claValoresCampoProducto
    Set proclaValoresCampoProducto.proConexion = Me.proConexion
    
    Set procolParametroProducto = New colParametroProducto
    Set procolParametroProducto.proConexion = Me.proConexion
    
    Set procolValordatos = New colValordatos
    Set procolValordatos.proConexion = Me.proConexion
    
    Set procolValoresCampoProducto = New colValoresCampoProducto
    Set procolValoresCampoProducto.proConexion = Me.proConexion
    'header
    Set proParametrosProducto = New colParametroProducto
    Set proParametrosProducto.proConexion = Me.proConexion
    Call SubFLlenarComboProductos
    Call AjustarvaloresEdicionCliente
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Sub LlenarListaDefinicion(Index As Integer, Optional proCampo As String = "")
    On Error GoTo ErrorManager
    Dim procolParametroProducto As colParametroProducto
    Set procolParametroProducto = New colParametroProducto
    Set procolParametroProducto.proConexion = Me.proConexion
    'limpiar las listas a derecha
    Dim byListas As Byte
        For byListas = Index To 2
            Me.lstDefinicion(byListas).Clear
        Next
    If cboCodigoProducto.ListIndex > -1 Then
        procolParametroProducto.proProductNumber = Trim(cboCodigoProducto.List(cboCodigoProducto.ListIndex))
        If Index = 0 Then
            procolParametroProducto.proCampoPadre = ""
        Else
            procolParametroProducto.proCampoPadre = Trim(Me.lstDefinicion(Index - 1).ListCodigo(lstDefinicion(Index - 1).ListIndex))
        End If
        If procolParametroProducto.metConsultarxProductoyCampo Then
            Dim i As Integer
            For i = 1 To procolParametroProducto.Count
                lstDefinicion(Index).AddItem procolParametroProducto.Item(i).proEtiqueta & " - " & Trim(procolParametroProducto.Item(i).proCampo) & " - " & DescripcionTipo(procolParametroProducto.Item(i).proTipo), procolParametroProducto.Item(i).proCampo, procolParametroProducto.Item(i).proTipo
            Next
        Else
            MsgBox "Error al consultar los Parametrosxproducto.", vbCritical, App.Title
        End If
        Set Me.procolParametroProducto = procolParametroProducto
    End If

Exit Sub
ErrorManager:
    SubGMuestraError

End Sub

Function DescripcionTipo(cTipo As String) As String
   On Error GoTo ErrorManager

    Select Case cTipo
            Case "T"
                DescripcionTipo = "Texto"
            Case "F"
                DescripcionTipo = "Fecha"
            Case "L"
                DescripcionTipo = "Lista"
            Case "B"
                DescripcionTipo = "Booleano"
    End Select

      Exit Function
ErrorManager:
    SubGMuestraError
End Function

Sub LlenarListaDatos(Index As Integer, Optional proCampo As String = "")
    On Error GoTo ErrorManager
    'limpiar las listas a derecha
    Dim byListas As Byte
    
    For byListas = Index To 2
        Me.lstDatos(byListas).Clear
    Next
    
    If lstDefinicion(Index).ListIndex <> -1 Or Index = 0 Then
        
        proCampo = Trim(lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex))
        procolValoresCampoProducto.proCampo = proCampo
        procolValoresCampoProducto.proProductNumber = Trim(cboCodigoProducto.List(cboCodigoProducto.ListIndex))
        If Index = 0 Then
            procolValoresCampoProducto.proValorIdPadre = 0
        Else
            procolValoresCampoProducto.proValorIdPadre = Val(Me.lstDatos(Index - 1).ListCodigo(lstDatos(Index - 1).ListIndex))
        End If
        If procolValoresCampoProducto.MetConsultarValoresxProducto Then
            Dim i As Integer
            For i = 1 To procolValoresCampoProducto.Count
                lstDatos(Index).AddItem procolValoresCampoProducto.Item(i).proValorDesc, _
                                                                      procolValoresCampoProducto.Item(i).proValorId
            Next
        Else
            MsgBox "Error al consultar.", vbCritical, App.Title
        End If
    End If

Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cboCodigoProducto_Click()
    On Error GoTo ErrorManager
    
    proProductNumber = Me.cboCodigoProducto.List(cboCodigoProducto.ListIndex)
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cboNombreProducto_Click()
    Dim i As Byte
        On Error GoTo ErrorManager
    If cboNombreProducto.ListIndex > -1 Then
        Me.cboCodigoProducto.ListIndex = Me.cboNombreProducto.ListIndex
        For i = 0 To 2
            Me.lstDatos(i).Clear
            Me.lstDefinicion(i).Clear
            lstDefinicion(i).MostrarBotonesV = True
            lstDatos(i).MostrarBotonesV = False
            lstDatos(i).MostrarBotonesH = False
        Next
        lstDefinicion(0).MostrarBotonesH = True
        Call LlenarListaDefinicion(0)
    Else
        lstDefinicion(0).MostrarBotonesV = False
        lstDefinicion(0).MostrarBotonesH = False
    End If
        
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmBuscarCampos_Click()

    On Error GoTo ErrorManager
    
    Screen.MousePointer = 11
    
    If Me.cboNombreProducto.Text = "" Or Me.cboNombreProducto.ListIndex = -1 Then
        MsgBox "Debe seleccionar el producto a buscar.", vbInformation, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Set Me.proParametrosProducto = Nothing
    Set Me.proParametrosProducto = New colParametroProducto
    Set Me.proParametrosProducto.proConexion = Me.proConexion
    
    Me.proParametrosProducto.proProductNumber = Me.cboCodigoProducto.Text
    
    If Me.proParametrosProducto.metConsultarxProductoyCampo Then
        
    Else
        MsgBox "Error al consultar los campos x producto.", vbCritical, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
ErrorManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub cmdGenerar_Click()
   On Error GoTo ErrorManager

    Set frmGeneraSubredIP.proConexion = Me.proConexion
    Set frmGeneraSubredIP.proclaValordatos = Me.proclaValordatos
    Set frmGeneraSubredIP.proclaValoresCampoProducto = Me.proclaValoresCampoProducto
    Set frmGeneraSubredIP.procolValordatos = Me.procolValordatos
    Set frmGeneraSubredIP.procolValoresCampoProducto = Me.procolValoresCampoProducto
    frmGeneraSubredIP.proCampoSubred = sCampoSubred
    frmGeneraSubredIP.proCampoIP = sCampoDireccionIP
    frmGeneraSubredIP.proIdPadre = lstDatos(0).ListCodigo(lstDatos(0).ListIndex)
    frmGeneraSubredIP.proProductNumber = Me.proProductNumber
    frmGeneraSubredIP.Show 1
    Call LlenarListaDatos(1, lstDatos(1).ListCodigo(lstDatos(1).ListIndex))

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorManager
    If Not bIniciada Then
        Call IniciarClases
        bIniciada = True
        Me.lstDefinicion(0).Texto = "Primer nivel"
        Me.lstDefinicion(1).Texto = "Segundo nivel"
        Me.lstDefinicion(2).Texto = "Tercer nivel"
        cboNombreProducto.Enabled = Not proEdicionCliente
    End If
    Screen.MousePointer = vbDefault

Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    bIniciada = False
End Sub

Private Sub lstDatos_btnAdicionarClick(Index As Integer)

On Error GoTo ErrorManager
If lstDefinicion(Index).ListIndex > -1 Then
    Set frmValor.proConexion = Me.proConexion
    frmValor.proProductNumber = Me.proProductNumber
    frmValor.proPermitirInsertar = True
    frmValor.proCampo = lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex)
    frmValor.ProOrden = lstDatos(Index).ListCount
    If Index = 0 Then
        frmValor.proValorIdPadre = 0
    Else
        frmValor.proValorIdPadre = lstDatos(Index - 1).ListCodigo(lstDatos(Index - 1).ListIndex)
    End If
    frmValor.Show 1
    Call LlenarListaDatos(Index, lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex))
Else
    MsgBox "Seleccione un parámetro"
End If
Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDatos_btnEliminarClick(Index As Integer)
'Eliminar Datos
On Error GoTo ErrorManager
    Dim i As Integer
    If lstDatos(Index).ListIndex > -1 Then
        If MsgBox("Esta acción eliminará  los valores de [" & Trim(lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex)) & "]  . Acepta?", vbYesNo, "Confirmación") = vbYes Then
                proclaValoresCampoProducto.proProductNumber = Me.proProductNumber
                proclaValoresCampoProducto.proCampo = Trim(lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex))
                proclaValoresCampoProducto.proValorId = Me.proValorId
                If Index > 0 Then
                    proclaValoresCampoProducto.proValorIdPadre = Trim(lstDatos(Index - 1).ListCodigo(lstDatos(Index - 1).ListIndex))
                Else
                     proclaValoresCampoProducto.proValorIdPadre = Trim(lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex))
                End If
                If proclaValoresCampoProducto.MetExistenRelaciones Then
                    MsgBox "No es posible eliminar. Existen hijos en CT_DatosProducto"
                Else
                    Call procolValoresCampoProducto.MetConsultarValoresxProducto
                    For i = 1 To procolValoresCampoProducto.Count
                        If procolValoresCampoProducto.Item(i).proValorIdPadre = proclaValoresCampoProducto.proValorId And procolValoresCampoProducto.Item(i).proValorId <> proclaValoresCampoProducto.proValorId Then
                            MsgBox "No es posible eliminar el valor. Es padre de almenos un valor [" & procolValoresCampoProducto.Item(i).proValorDesc & "]"
                            Exit Sub
                        End If
                    Next
                    Call proclaValoresCampoProducto.MetEliminar
                End If
        End If
        Call LlenarListaDatos(Index, lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex))
        
    End If
        
Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDatos_Click(Index As Integer)
    On Error GoTo ErrorManager
    Dim byListas As Integer
If Index < 2 Then
    For byListas = Index + 1 To 2
        Me.lstDatos(byListas).MostrarBotonesH = False
    Next
End If

If lstDatos(Index).ListIndex > -1 Then
    
    If Index >= 0 Then
        If Index < 2 Then
            Call LlenarListaDatos(Index + 1, lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex))
            lstDatos(Index + 1).MostrarBotonesH = True
        End If
        proValorId = lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex)
    End If
Else
    If Index < 2 Then lstDatos(Index + 1).MostrarBotonesH = False
End If

If (cboCodigoProducto.List(cboCodigoProducto.ListIndex) = TipoTelefoniaLocal Or cboCodigoProducto.List(cboCodigoProducto.ListIndex) = TipoTelefoniaNacional) And Trim(lstDefinicion(1).ListCodigo(lstDefinicion(1).ListIndex)) = sCampoSubred Then
    Me.cmdGenerar.Visible = Not proEdicionCliente
Else
    Me.cmdGenerar.Visible = False
End If

Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDatos_DblClick(Index As Integer)
        On Error GoTo ErrorManager
    Set frmValor.proConexion = Me.proConexion
    frmValor.proPermitirInsertar = False
    If Index = 0 Then
        frmValor.proValorIdPadre = 0
    Else
        frmValor.proValorIdPadre = lstDatos(Index - 1).ListCodigo(lstDatos(Index - 1).ListIndex)
    End If
    frmValor.proValorId = lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex)
    frmValor.proProductNumber = Me.proProductNumber
    frmValor.proCampo = lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex)
    frmValor.Show 1

Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDefinicion_CambioOrden(Index As Integer)
    Dim i As Integer, byListas As Integer
    On Error GoTo ErrorManager
    For i = 0 To lstDefinicion(Index).ListCount - 1
        proclaParametroProducto.proCampo = lstDefinicion(Index).ListCodigo(i)
        If Index = 0 Then
            proclaParametroProducto.proCampoPadre = ""
        Else
            proclaParametroProducto.proCampoPadre = lstDefinicion(Index - 1).ListCodigo(lstDefinicion(Index - 1).ListIndex)
        End If
        proclaParametroProducto.proProductNumber = proProductNumber
        proclaParametroProducto.ProOrden = i + 1
        proclaParametroProducto.MetActualizarOrden
    Next
    
        Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDefinicion_CambioOrdenAZ(Index As Integer)
    Dim byListas As Integer
   On Error GoTo ErrorManager

    Call lstDefinicion_CambioOrden(Index)
    If Index < 2 Then
        If Index = 0 Then
            Me.lstDatos(0).Clear
            lstDefinicion(0).ListIndex = -1
        End If
        For byListas = Index + 1 To 2
            Me.lstDefinicion(byListas).MostrarBotonesH = False
            Me.lstDatos(byListas).MostrarBotonesV = False
            Me.lstDefinicion(byListas).Clear
            Me.lstDatos(byListas).Clear
        Next
    End If

      Exit Sub
ErrorManager:
    SubGMuestraError

End Sub

Private Sub lstDefinicion_LostFocus(Index As Integer)
On Error GoTo ErrorManager
    
    If Index < 2 Then
        If lstDefinicion(Index).ListCount = 0 Then lstDefinicion(Index + 1).MostrarBotonesH = False
    End If
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDatos_LostFocus(Index As Integer)
On Error GoTo ErrorManager
    If Index < 2 Then
        If lstDatos(Index).ListCount = 0 Then lstDatos(Index + 1).MostrarBotonesH = False
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDefinicion_btnAdicionarClick(Index As Integer)
On Error GoTo ErrorManager
    Screen.MousePointer = vbHourglass
    Set frmCamposProducto.proConexion = Me.proConexion
    frmCamposProducto.proProductNumber = Me.proProductNumber
    frmCamposProducto.ProOrden = lstDefinicion(Index).ListCount
    If Index = 0 Then
        frmCamposProducto.proCampoPadre = 0
    Else
        frmCamposProducto.proCampoPadre = lstDefinicion(Index - 1).ListCodigo(lstDefinicion(Index - 1).ListIndex)
    End If
    Call frmCamposProducto.PriductListIndex(cboNombreProducto.ListIndex)
    frmCamposProducto.SubFLlenarComboCampos ("N")
    Call frmCamposProducto.Insertar(True)
    If Index > 0 Then
        frmCamposProducto.cboTipo.ListIndex = 2 'Hardcode tipo lista
        frmCamposProducto.cboCodigoTipo.ListIndex = 2
        frmCamposProducto.cboTipo.Enabled = False
    End If
    frmCamposProducto.Show 1
    Call LlenarListaDefinicion(Index, lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex))
    Call LlenarListaDatos(Index, lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex))
    lstDatos(Index).MostrarBotonesV = False
    lstDatos(Index).Texto = ""
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDefinicion_btnEliminarClick(Index As Integer)
    'Eliminar definición
    Dim i As Integer
    On Error GoTo ErrorManager
    
    If lstDefinicion(Index).ListIndex > -1 Then
        If MsgBox("Esta acción eliminará la definición y los valores de [" & Trim(lstDefinicion(Index).ListDescripcion(lstDefinicion(Index).ListIndex)) & "] de esta definición . Acepta?", vbYesNo, "Confirmación") = vbYes Then
                proclaParametroProducto.proProductNumber = Me.proProductNumber
                proclaParametroProducto.proCampo = lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex)
                If proclaParametroProducto.MetExistenRelaciones Then
                    MsgBox "No es posible eliminar. Existen relaciones en ct_detalledatosproducto para el campo " & proclaParametroProducto.proCampo
                Else
                    Call procolParametroProducto.metConsultarxProducto
                    For i = 1 To procolParametroProducto.Count
                        If procolParametroProducto.Item(i).proCampoPadre = proclaParametroProducto.proCampo And procolParametroProducto.Item(i).proCampo <> proclaParametroProducto.proCampo Then
                            MsgBox "No es posible eliminar el parametro [" & Trim(lstDefinicion(Index).ListDescripcion(lstDefinicion(Index).ListIndex)) & "]" & " . Es padre de otros parámetros"
                            Exit Sub
                        End If
                    Next
                    Call proclaParametroProducto.MetEliminarValoresCampo
                    Call proclaParametroProducto.MetEliminar
                End If
        End If
        Call LlenarListaDefinicion(Index, lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex))
        Call LlenarListaDatos(Index, lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex))
    End If
    Exit Sub
    lstDatos(Index).MostrarBotonesV = False
    lstDatos(Index).Texto = ""
ErrorManager:
    SubGMuestraError

End Sub

Private Sub lstDefinicion_Click(Index As Integer)

    On Error GoTo ErrorManager
    Dim byListas As Integer
    If Index < 2 Then
        For byListas = Index + 1 To 2
            Me.lstDefinicion(byListas).MostrarBotonesH = False
            lstDefinicion(byListas).Clear
            lstDatos(byListas).Clear
            Me.lstDatos(byListas).MostrarBotonesV = False
            Me.lstDatos(byListas).Texto = ""
        Next
    End If
    If lstDefinicion(Index).ListIndex > -1 Then
        If lstDefinicion(Index).ListCampoAdicional(lstDefinicion(Index).ListIndex) = "L" Then
            lstDatos(Index).MostrarBotonesV = True
            If Index = 0 Then
                lstDatos(Index).MostrarBotonesH = True
                Call LlenarListaDatos(Index)
            End If
            If Index < 2 Then
                lstDefinicion(Index + 1).MostrarBotonesH = True
                Call LlenarListaDefinicion(Index + 1, lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex))
            End If
            If Index > 0 Then
                If lstDatos(Index - 1).ListIndex > -1 Then
                    Call LlenarListaDatos(Index, lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex))
                Else
                    lstDatos(Index).MostrarBotonesH = False
                End If
            End If
            
        Else
                lstDatos(0).MostrarBotonesV = False
                lstDatos(Index).MostrarBotonesV = False
                For byListas = Index To 2
                    Me.lstDatos(byListas).Clear
                Next
        End If
        Me.lstDatos(Index).Texto = Me.lstDefinicion(Index).ListDescripcion(lstDefinicion(Index).ListIndex)
    End If
    If (cboCodigoProducto.List(cboCodigoProducto.ListIndex) = TipoTelefoniaLocal Or cboCodigoProducto.List(cboCodigoProducto.ListIndex) = TipoTelefoniaNacional) And Trim(lstDefinicion(1).ListCodigo(lstDefinicion(1).ListIndex)) = sCampoSubred Then
        Me.cmdGenerar.Visible = True
    Else
        Me.cmdGenerar.Visible = False
    End If
    
    Exit Sub
ErrorManager:
End Sub

Private Sub lstDefinicion_DblClick(Index As Integer)

On Error GoTo ErrorManager
    Set frmCamposProducto.proConexion = Me.proConexion
    frmCamposProducto.proProductNumber = Me.proProductNumber
    frmCamposProducto.proCampo = Trim(lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex))
    Call frmCamposProducto.PriductListIndex(cboNombreProducto.ListIndex)
    frmCamposProducto.SubFLlenarComboCampos ("N")
    Call frmCamposProducto.BuscarRegistroEnGrid
    Call frmCamposProducto.grdCampos_Click
    Call frmCamposProducto.Insertar(False)
    Call frmCamposProducto.BuscarCampo
    frmCamposProducto.Show 1
    If Index < 2 Then lstDefinicion(Index + 1).MostrarBotonesH = False
    lstDatos(Index).MostrarBotonesV = False
    lstDatos(Index).Texto = ""
    Call LlenarListaDefinicion(Index, lstDefinicion(Index).ListCodigo(lstDefinicion(Index).ListIndex))
    Call LlenarListaDatos(Index, lstDatos(Index).ListCodigo(lstDatos(Index).ListIndex))
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

