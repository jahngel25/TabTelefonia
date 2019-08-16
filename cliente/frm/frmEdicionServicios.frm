VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEdicionServicios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edición de Servicios Suplementarios"
   ClientHeight    =   7215
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   17160
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ckbTodos 
      Caption         =   "Aplicar a Todos los Números"
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   6840
      Width           =   2505
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   315
      Left            =   15870
      TabIndex        =   1
      Top             =   6810
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSFlexGridLib.MSFlexGrid grdServicios 
      Height          =   6765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17325
      _ExtentX        =   30559
      _ExtentY        =   11933
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   5
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSForms.Label lbltexto 
      Height          =   405
      Left            =   4740
      TabIndex        =   5
      Top             =   6810
      Width           =   10965
      Caption         =   $"frmEdicionServicios.frx":0000
      Size            =   "19341;714"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblModificar 
      BackColor       =   &H00F9FCE7&
      BorderStyle     =   1  'Fixed Single
      Height          =   165
      Left            =   60
      TabIndex        =   3
      Top             =   6900
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Celdas modificadas"
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
      Left            =   330
      TabIndex        =   2
      Top             =   6870
      Width           =   1410
   End
End
Attribute VB_Name = "frmEdicionServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmEdicionServiciosFormulario
' Fecha  : 12/10/2004 15:38
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : Administrar la edición de servicios suplementarios en telefonía pública
'---------------------------------------------------------------------------------------
Option Explicit

Public proDatosProducto As claDatosProducto

'Encabezado de servicios suplementarios
Public procolServiciosSup As EDCAdminVoz.colServiciosSup

'Para obtener los numeros
Public procolDatosProductoNumero As colDatosProductoNumero
Public procolNovedadNumero As colNovedadNumero

'Para adicionar novedad al numero al detalle
Public proclaNovedadNumero As claNovedadNumero

'Para los servicios actuales
Public proclaServiciosxNumero As claServiciosxNumero
Public procolServiciosxNumero As colServiciosxNumero

'para los servicios nuevos
Public proclaServiciosxReserva As claServiciosxReserva
Public procolServiciosxReserva As colserviciosxreserva

'para los valores de los servicios suplementarios Nuevos
Public varClaValorServicioxNumero As claValorServicioxnumero
Public varClaNovedadvalorServicioxNumero As claNovedadValorServicioxNumero

Public procolValorServicioxReserva As colNovedadValorServicioxNumero

Public procolValorServicioxNumero  As colValorServicioxnumero


'Conexion a la base de datos
Public proConexion As ADODB.Connection

Public proEsNovedad As Integer
'EsNovedad = 1
'NoesNovedad = 0

Public proAccionValor As Integer
'=1 si se adiciono
'=2 si elimino un valor
'=0 si no se realizo ninguna accion

Public proAplicaraTodos As Integer
'=1 si aplicar a todos los números
'=0 no aplicar a todos los números


'valores de cada celda checked unchecked
Const strChecked = "þ"
Const strUnChecked = "q"
Const strColorAzul = &HF9FCE7
Const iCVienedeNovedadNumero = 1
Const iCVienedeDatosProNum = 0
Dim ColumnaSeleccionada As Integer

Private Sub ckbTodos_Click()
On Error GoTo ErrorManager
    If ckbTodos.Value Then
        proAplicaraTodos = 1
    Else
        proAplicaraTodos = 0
    End If
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
    Dim iCol As Integer
    Dim iRow As Integer
    Dim iColAnt As Integer
    Dim iRowAnt As Integer
    
    Dim iColSelec As Integer
    Dim iRow1 As Integer
    Dim sValor As String
   On Error GoTo ErrorManager
   Screen.MousePointer = 11
   
    If Me.proAplicaraTodos = 1 Then
        'DEBE UBICAR LA COLUNMA SELECCIONADA Y RECORRER TODAS SUS FILAS
        With grdServicios
            .Redraw = False
            For iCol = 5 To .Cols - 1
                For iRow = 2 To .Rows - 1
                    .Row = iRow
                    .Col = iCol
                    If .CellBackColor = strColorAzul Then 'UBICA CUAL ES LA COLUMNA Y FILA SELECCIONADA
                        'RECORRE TODAS LA FILAS DE ESA COLUMNA
                        iColSelec = iCol
                        .Col = iColSelec
                        sValor = .TextMatrix(iRow, iColSelec)
                        For iRow1 = 2 To .Rows - 1
                            'Almacena en las otras tablas
                            If Not GuardarCelda(iRow1, iColSelec, sValor) Then
                                MsgBox "No fue posible guardar el valor para el número " & grdServicios.TextMatrix(iRow, 4) & " en el servicio [" & grdServicios.TextMatrix(1, iCol) & "}"
                            Else
                                iColAnt = iColSelec
                                iRowAnt = iRow1
                            End If
                            'Se almacena en la tabla de valores servicio por numero (cuando es de tipocheckbox)
                            If Not GuardarValorServicio(iRow1, iColSelec, sValor) Then
                                MsgBox "No fue posible guardar el valor para el número " & grdServicios.TextMatrix(iRow, 4) & " en el servicio [" & grdServicios.TextMatrix(1, iCol) & "}"
                            End If
                        Next iRow1
                        Exit For
                    End If
                Next iRow
            Next iCol
            .Redraw = True
            .Col = iColAnt
            .Row = iRowAnt
        End With
        'Debe Actualizar la Grilla
        Call Form_Activate
        proAplicaraTodos = 0
        Me.ckbTodos.Value = False
        
    Else
        With grdServicios
            .AllowBigSelection = False
            .Redraw = False
            For iRow = 2 To .Rows - 1
                For iCol = 5 To .Cols - 1
                    .Row = iRow
                    .Col = iCol
                    If .CellBackColor = strColorAzul Then 'UBICA CUAL ES LA COLUMNA Y FILA SELECCIONADA
                            'Almacena en las otras tablas
                            If Not GuardarCelda(iRow, iCol, .TextMatrix(iRow, iCol)) Then
                                MsgBox "No fue posible guardar el valor para el número " & grdServicios.TextMatrix(iRow, 4) & " en el servicio [" & grdServicios.TextMatrix(1, iCol) & "}"
                            Else
                                iColAnt = iCol
                                iRowAnt = iRow
                            End If
                            'Se almacena en la tabla de valores servicio por numero (cuando es de tipocheckbox)
                            If Not GuardarValorServicio(iRow, iCol, .TextMatrix(iRow, iCol)) Then
                                MsgBox "No fue posible guardar el valor para el número " & grdServicios.TextMatrix(iRow, 4) & " en el servicio [" & grdServicios.TextMatrix(1, iCol) & "}"
                            End If
                            .CellBackColor = 0
                    End If
                Next iCol
            Next iRow
            .Redraw = True
            .Col = iColAnt
            .Row = iRowAnt
        End With
        
    End If
    
    'MsgBox "Se guardaron exitosamente los cambios"
    Me.ckbTodos.SetFocus
    Screen.MousePointer = 0
    'Unload Me
    Exit Sub
ErrorManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub

Function GuardarCelda(iRow As Integer, iCol As Integer, valor As String) As Boolean
   On Error GoTo ErrorManager
    
    proclaServiciosxReserva.proServicioSuplementarioId = Me.grdServicios.TextMatrix(iRow, 0)
    'proclaServiciosxReserva.proServiciosuplementarioId = procolServiciosSup.Item(iCol - 4).proiServicioSuplementarioId
    
    If valor = strChecked Then
        'Adicionar
        'Primero debe validar si ese número ya existe en la CT_DATOSPRODUCTONUMERO
        procolDatosProductoNumero.proRegionCode = grdServicios.TextMatrix(iRow, 3)
        procolDatosProductoNumero.proNumero = grdServicios.TextMatrix(iRow, 4)
        procolDatosProductoNumero.MetConsultarExistenciaNumero
        If procolDatosProductoNumero.Count > 0 Then
            'Quiere decir que ya existe como número asignado al servicio del cliente
            'Debemos Valiar que exista en la Tabla CT_NOVEDADNUMERO
            procolNovedadNumero.proRegionCode = grdServicios.TextMatrix(iRow, 3)
            procolNovedadNumero.proNumero = grdServicios.TextMatrix(iRow, 4)
            procolNovedadNumero.proDatosProductoId = procolDatosProductoNumero.Item(1).proDatosProductoId
            procolNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
            If Not procolNovedadNumero.MetConsultarxServicio Then
            'Quiere decir que estan realizando algo de servcios suplementarios unicamente con este número
            'Como no existe debemos insertarlo en la tabla pero de tipo =2 Modifcación
            'se debe adicionar  como ct_NovedadNumero
                proclaNovedadNumero.proRegionCode = Me.grdServicios.TextMatrix(iRow, 3)
                proclaNovedadNumero.proNumero = Me.grdServicios.TextMatrix(iRow, 4)
                proclaNovedadNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
                proclaNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
                proclaNovedadNumero.proTipoNovedadId = "2"
                proclaNovedadNumero.proFechaReserva = Format(Now, "mm/dd/yyyy hh:mm:ss")
                proclaNovedadNumero.proFechaLiberacion = ""
                If Not proclaNovedadNumero.FunGInsertar Then
                        MsgBox "No se adicionó el registro en NovedadNumero"
                        Exit Function
                End If
                'Lo agrega a la coleccion
                Call proDatosProducto.MetAgregarNovedadNumeroPublico(proclaNovedadNumero)
                proclaServiciosxReserva.proNovedadNumeroId = proclaNovedadNumero.proNovedadNumeroId
                proclaServiciosxReserva.proTipoNovedadId = proclaNovedadNumero.proTipoNovedadId
            Else
                proclaServiciosxReserva.proNovedadNumeroId = procolNovedadNumero.proNovedadNumeroId
                proclaServiciosxReserva.proTipoNovedadId = procolNovedadNumero.proTipoNovedadId
            End If
            proclaServiciosxReserva.proTipoNovedadId = "1"
            proclaServiciosxReserva.proServicioSuplementarioId = grdServicios.TextMatrix(0, iCol)
            proclaServiciosxReserva.FunGInsertar
        Else
            'Quiere decir que no existe como número agregado del cliente
            'Debemos validar entonces que tiene que existir en la tabla de Ct_novedadNumero
            procolNovedadNumero.proRegionCode = grdServicios.TextMatrix(iRow, 3)
            procolNovedadNumero.proNumero = grdServicios.TextMatrix(iRow, 4)
            procolNovedadNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
            procolNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
            If Not procolNovedadNumero.MetConsultarxServicio Then
                proclaNovedadNumero.proRegionCode = Me.grdServicios.TextMatrix(iRow, 3)
                proclaNovedadNumero.proNumero = Me.grdServicios.TextMatrix(iRow, 4)
                proclaNovedadNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
                proclaNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
                proclaNovedadNumero.proTipoNovedadId = "2"
                proclaNovedadNumero.proFechaReserva = Format(Now, "mm/dd/yyyy hh:mm:ss")
                proclaNovedadNumero.proFechaLiberacion = ""
                If Not proclaNovedadNumero.FunGInsertar Then
                        MsgBox "No se adicionó el registro en NovedadNumero"
                        Exit Function
                End If
                'Lo agrega a la coleccion
                Call proDatosProducto.MetAgregarNovedadNumeroPublico(proclaNovedadNumero)
                proclaServiciosxReserva.proNovedadNumeroId = proclaNovedadNumero.proNovedadNumeroId
                proclaServiciosxReserva.proTipoNovedadId = proclaNovedadNumero.proTipoNovedadId
            Else
                'Si existe la Novedad
                proclaServiciosxReserva.proNovedadNumeroId = procolNovedadNumero.proNovedadNumeroId
                proclaServiciosxReserva.proTipoNovedadId = procolNovedadNumero.proTipoNovedadId
            End If
            proclaServiciosxReserva.proTipoNovedadId = "1"
            proclaServiciosxReserva.proServicioSuplementarioId = grdServicios.TextMatrix(0, iCol)
            proclaServiciosxReserva.FunGInsertar
        End If
    Else ' Del Check No esta Seleccionado
        'Eliminacion servicio
        'Primero debe validar si ese número ya existe en la CT_DATOSPRODUCTONUMERO
        procolDatosProductoNumero.proRegionCode = grdServicios.TextMatrix(iRow, 3)
        procolDatosProductoNumero.proNumero = grdServicios.TextMatrix(iRow, 4)
        procolDatosProductoNumero.MetConsultarExistenciaNumero
        If procolDatosProductoNumero.Count > 0 Then
            'Quiere decir que ya existe como número asignado al servicio del cliente
            'Debemos Valiar que exista en la Tabla CT_NOVEDADNUMERO
            procolNovedadNumero.proRegionCode = grdServicios.TextMatrix(iRow, 3)
            procolNovedadNumero.proNumero = grdServicios.TextMatrix(iRow, 4)
            procolNovedadNumero.proDatosProductoId = procolDatosProductoNumero.Item(1).proDatosProductoId
            procolNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
            If Not procolNovedadNumero.MetConsultarxServicio Then
            'Quiere decir que estan realizando algo de servcios suplementarios unicamente con este número
            'Como no existe debemos insertarlo en la tabla pero de tipo =2 Modifcación
            'se debe adicionar  como ct_NovedadNumero
                proclaNovedadNumero.proRegionCode = Me.grdServicios.TextMatrix(iRow, 3)
                proclaNovedadNumero.proNumero = Me.grdServicios.TextMatrix(iRow, 4)
                proclaNovedadNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
                proclaNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
                proclaNovedadNumero.proTipoNovedadId = "2"
                proclaNovedadNumero.proFechaReserva = Format(Now, "mm/dd/yyyy hh:mm:ss")
                proclaNovedadNumero.proFechaLiberacion = ""
                If Not proclaNovedadNumero.FunGInsertar Then
                        MsgBox "No se adicionó el registro en NovedadNumero"
                        Exit Function
                End If
                'Lo agrega a la coleccion
                Call proDatosProducto.MetAgregarNovedadNumeroPublico(proclaNovedadNumero)
                proclaServiciosxReserva.proNovedadNumeroId = proclaNovedadNumero.proNovedadNumeroId
                proclaServiciosxReserva.proTipoNovedadId = proclaNovedadNumero.proTipoNovedadId
                
            Else
                proclaServiciosxReserva.proNovedadNumeroId = procolNovedadNumero.proNovedadNumeroId
                proclaServiciosxReserva.proTipoNovedadId = procolNovedadNumero.proTipoNovedadId
            End If
                proclaServiciosxReserva.proTipoNovedadId = "3"
                proclaServiciosxReserva.proServicioSuplementarioId = grdServicios.TextMatrix(0, iCol)
                proclaServiciosxReserva.FunGInsertar
            
        Else
            'Quiere decir que no existe como número agregado del cliente
            'Debemos validar entonces que tiene que existir en la tabla de Ct_novedadNumero
            procolNovedadNumero.proRegionCode = grdServicios.TextMatrix(iRow, 3)
            procolNovedadNumero.proNumero = grdServicios.TextMatrix(iRow, 4)
            procolNovedadNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
            procolNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
            If Not procolNovedadNumero.MetConsultarxServicio Then
                proclaNovedadNumero.proRegionCode = Me.grdServicios.TextMatrix(iRow, 3)
                proclaNovedadNumero.proNumero = Me.grdServicios.TextMatrix(iRow, 4)
                proclaNovedadNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
                proclaNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
                proclaNovedadNumero.proTipoNovedadId = "1"
                proclaNovedadNumero.proFechaReserva = Format(Now, "mm/dd/yyyy hh:mm:ss")
                proclaNovedadNumero.proFechaLiberacion = ""
                If Not proclaNovedadNumero.FunGInsertar Then
                        MsgBox "No se adicionó el registro en NovedadNumero"
                        Exit Function
                End If
                'Lo agrega a la coleccion
                Call proDatosProducto.MetAgregarNovedadNumeroPublico(proclaNovedadNumero)
                proclaServiciosxReserva.proNovedadNumeroId = proclaNovedadNumero.proNovedadNumeroId
                proclaServiciosxReserva.proTipoNovedadId = proclaNovedadNumero.proTipoNovedadId
            Else
                proclaServiciosxReserva.proNovedadNumeroId = procolNovedadNumero.proNovedadNumeroId
                proclaServiciosxReserva.proTipoNovedadId = procolNovedadNumero.proTipoNovedadId
            End If
            'Elimina los servcios
            proclaServiciosxReserva.proNovedadNumeroId = Me.grdServicios.TextMatrix(iRow, 2)
            proclaServiciosxReserva.proServicioSuplementarioId = Me.grdServicios.TextMatrix(0, iCol)
            proclaServiciosxReserva.FunGEliminar
        End If
    End If
    GuardarCelda = True
    Exit Function
ErrorManager:
    GuardarCelda = False
    SubGMuestraError
End Function


Function GuardarValorServicio(iRow As Integer, iCol As Integer, valor As String) As Boolean
   On Error GoTo ErrorManager
    
    If valor = strChecked Then
        'Inserta el valor
        'Quiere decir que debo almacenar la información en la tabla CT_NOVEDADVALORSERVICIOXNUMERO
        Set varClaNovedadvalorServicioxNumero = New claNovedadValorServicioxNumero
        Set varClaNovedadvalorServicioxNumero.proConexion = Me.proConexion
        varClaNovedadvalorServicioxNumero.proNovedadNumeroId = Me.proclaServiciosxReserva.proNovedadNumeroId
        varClaNovedadvalorServicioxNumero.proServicioSuplementario = Me.grdServicios.TextMatrix(0, iCol)
        varClaNovedadvalorServicioxNumero.proValor = "1"
        varClaNovedadvalorServicioxNumero.proNumero = Me.grdServicios.TextMatrix(iRow, 4)
        varClaNovedadvalorServicioxNumero.proRegion = Me.grdServicios.TextMatrix(iRow, 3)
        If Not varClaNovedadvalorServicioxNumero.FunGInsertar Then
            MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
        End If
        Set varClaNovedadvalorServicioxNumero = Nothing
        'Fin Insertar valor
    Else
        'Elimina el valor
        Set varClaNovedadvalorServicioxNumero = New claNovedadValorServicioxNumero
        
        Set varClaNovedadvalorServicioxNumero.proConexion = Me.proConexion
        varClaNovedadvalorServicioxNumero.proNovedadNumeroId = proclaServiciosxReserva.proNovedadNumeroId
        varClaNovedadvalorServicioxNumero.proServicioSuplementario = Me.grdServicios.TextMatrix(0, iCol)
        varClaNovedadvalorServicioxNumero.proNumero = Me.grdServicios.TextMatrix(iRow, 4)
        varClaNovedadvalorServicioxNumero.proRegion = Me.grdServicios.TextMatrix(iRow, 3)
        If Not varClaNovedadvalorServicioxNumero.FunGEliminar Then
            MsgBox "No fue posible eliminar el valor para el servicio suplementario.", vbInformation, App.Title
        End If
        Set varClaNovedadvalorServicioxNumero = Nothing
    End If
    GuardarValorServicio = True
    Exit Function
ErrorManager:
    GuardarValorServicio = False
    SubGMuestraError
End Function

Private Sub Form_Activate()
   On Error GoTo ErrorManager
    Screen.MousePointer = 11
    Call PintarGrid
    Call LlenarGrid
    Screen.MousePointer = 0
      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorManager

    Set proclaServiciosxReserva = New claServiciosxReserva
    Set proclaServiciosxReserva.proConexion = Me.proConexion
    
    Set procolServiciosxReserva = New colserviciosxreserva
    Set procolServiciosxReserva.proConexion = Me.proConexion
    
    Set proclaServiciosxNumero = New claServiciosxNumero
    Set proclaServiciosxNumero.proConexion = Me.proConexion
    
    Set procolDatosProductoNumero = New colDatosProductoNumero
    Set procolDatosProductoNumero.proConexion = Me.proConexion
    
    Set procolNovedadNumero = New colNovedadNumero
    Set procolNovedadNumero.proConexion = Me.proConexion
    
    Set proclaNovedadNumero = New claNovedadNumero
    Set proclaNovedadNumero.proConexion = Me.proConexion
    
    Set procolServiciosSup = New colServiciosSup
    Set procolServiciosSup.proConexion = Me.proConexion
    
    Set procolServiciosxNumero = New colServiciosxNumero
    Set procolServiciosxNumero.proConexion = Me.proConexion
    
    Set procolServiciosxNumero = New colServiciosxNumero
    Set procolServiciosxNumero.proConexion = Me.proConexion

    'para manejo de los valores de los servicios

    Set procolValorServicioxReserva = New colNovedadValorServicioxNumero
    Set procolValorServicioxReserva.proConexion = Me.proConexion
    
    Set procolValorServicioxNumero = New colValorServicioxnumero
    Set procolValorServicioxNumero.proConexion = Me.proConexion
    
    'Para validar existencia
    Set procolDatosProductoNumero = New colDatosProductoNumero
    Set procolDatosProductoNumero.proConexion = Me.proConexion


      Exit Sub
ErrorManager:
    SubGMuestraError

End Sub

Private Sub Form_Resize()
   On Error GoTo ErrorManager

    grdServicios.Width = Me.ScaleWidth
    grdServicios.Height = Me.ScaleHeight - 300
    Me.cmdGuardar.Top = Me.ScaleHeight - 400
    cmdGuardar.Left = Me.ScaleWidth - 1450
    Me.Label3.Top = Me.cmdGuardar.Top
    Me.lblModificar.Top = cmdGuardar.Top
    Me.ckbTodos.Top = Me.cmdGuardar.Top
    Me.lbltexto.Top = Me.cmdGuardar.Top

      Exit Sub
ErrorManager:
    'SubGMuestraError
End Sub

Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
    Dim sServicioSuplementarioId As String
    Dim sNombreServicioSuplementarioId As String
    Dim sLinea As String
    Dim sregion As String
    Dim sAcciones As String
    Dim iRow2 As Integer
    Dim sValor As String

    
    
   On Error GoTo ErrorManager
   
    'validamos que tipo de servicios es si es check sigue el flujo normal
    If procolServiciosSup.Item(iCol - 4).prochTipoServicio = "C" Then 'Es un Checkbox
        If proAplicaraTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
            If MsgBox("Desea aplicar estos cambios a todos los números con este servicio suplementario?", vbYesNo + vbInformation, App.Title) = vbNo Then
                Exit Sub
            End If
        End If

        With grdServicios
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
            Else
                .TextMatrix(iRow, iCol) = strUnChecked
            End If
            .TextMatrix(iRow, 0) = 1
            .Row = iRow
            .Col = iCol
            If .CellBackColor = strColorAzul Then
                .CellBackColor = vbWhite
            Else
                .CellBackColor = strColorAzul
            End If
            Call cmdGuardar_Click
        End With
    Else ' Es un Combo o un Texto
        sServicioSuplementarioId = procolServiciosSup.Item(iCol - 4).proiServicioSuplementarioId
        sNombreServicioSuplementarioId = procolServiciosSup.Item(iCol - 4).provchNombreServicio
        sLinea = grdServicios.TextMatrix(iRow, 4)
        sregion = grdServicios.TextMatrix(iRow, 3)
        
        Set frmValorServicioSup.proConexion = Me.proConexion
        frmValorServicioSup.proServicioSuplementario = sServicioSuplementarioId
        frmValorServicioSup.proNombreServicioSuplementario = sNombreServicioSuplementarioId
        frmValorServicioSup.proTelefono = sLinea
        frmValorServicioSup.proRegion = sregion
        frmValorServicioSup.proTodos = Me.proAplicaraTodos
        frmValorServicioSup.proTipoServicio = procolServiciosSup.Item(iCol - 4).prochTipoServicio
        frmValorServicioSup.proValor = grdServicios.TextMatrix(iRow, iCol)
        frmValorServicioSup.proNovedadId = Me.grdServicios.TextMatrix(iRow, 2)
        'If procolNovedadNumero.Count > 0 Then
        '    frmValorServicioSup.proNovedadId = Me.grdServicios.TextMatrix(iRow, 2)
        'Else
        '    frmValorServicioSup.proNovedadId = 0
        'End If
        frmValorServicioSup.Show (1)
        If Me.proAccionValor = 1 Or Me.proAccionValor = 2 Then
            If Me.proAccionValor = 1 Then 'Guardar
                sAcciones = strChecked
            ElseIf Me.proAccionValor = 2 Then 'eliminar
                sAcciones = strUnChecked
            End If
            If Me.proAplicaraTodos = 1 Then
                sValor = sAcciones
                'debe ser aplicado a todas las filas de la columna seleccionada
                For iRow2 = 2 To Me.grdServicios.Rows - 1
                    'Almacena en las otras tablas
                    If Not GuardarCelda(iRow2, iCol, sValor) Then
                        MsgBox "No fue posible guardar el valor para el número " & grdServicios.TextMatrix(iRow2, 4) & " en el servicio [" & grdServicios.TextMatrix(1, iCol) & "}"
                    End If
                Next iRow2
            Else
                'Guarda en las tablas correspondientes
                If Not GuardarCelda(iRow, iCol, sAcciones) Then
                    MsgBox "No fue posible guardar el valor para el número " & grdServicios.TextMatrix(iRow, 4) & " en el servicio [" & grdServicios.TextMatrix(1, iCol) & "}"
                End If
            End If
            'Debe Actualizar la Grilla
            Call Form_Activate
            Me.proAccionValor = 0
            proAplicaraTodos = 0
            Me.ckbTodos.Value = False
        End If
        
    End If
        
      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub grdServicios_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrorManager

    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
        With grdServicios
            Call TriggerCheckbox(.Row, .Col)
        End With
    End If

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub grdServicios_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo ErrorManager
With grdServicios
If .MouseCol > 4 Then
        If Button = 1 And .MouseRow > 1 Then
                Call TriggerCheckbox(.MouseRow, .MouseCol)
        'Else
        '    grdServicios.Col = .MouseCol
        '    grdServicios.ColSel = .MouseCol
        End If
    End If
    End With
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Sub PintarGrid()
    Dim iContador As Long
   On Error GoTo ErrorManager

        procolServiciosSup.prochProductNumber = proDatosProducto.proProductNumber
        If Not procolServiciosSup.FunGConsulta Then
            MsgBox "No se pudieron consultar los servicios suplementarios"
            Exit Sub
        End If
        'Limpia la grilla
        Me.grdServicios.Rows = 1
        Me.grdServicios.Refresh
        With grdServicios
            .Rows = 2
            .Cols = procolServiciosSup.Count + 5
            .TextMatrix(1, 2) = "ID"
            .TextMatrix(1, 3) = "Region"
            .TextMatrix(1, 4) = "N°"
            For iContador = 5 To procolServiciosSup.Count + 4
                .TextMatrix(0, iContador) = procolServiciosSup.Item(iContador - 4).proiServicioSuplementarioId
                .TextMatrix(1, iContador) = procolServiciosSup.Item(iContador - 4).provchNombreServicio
                .ColWidth(0) = 1000
                .Col = iContador
                .Row = 1
                .CellFontName = "Arial Narrow"
                .ColWidth(iContador) = 2750
                 
            Next
            'ocultar columnas y filas propuestas como  ocultas
            .RowHeight(0) = 30
            .ColWidth(0) = 0
            .ColWidth(1) = 0
            .ColWidth(2) = 0
        End With

      Exit Sub
ErrorManager:
    SubGMuestraError
        
End Sub

Sub LlenarGrid()
        Dim TotalFilas As Long
        Dim iContadorPro As Long
        Dim iContadorNov As Long
        Dim bEsta As Boolean
        Dim iCol As Integer
        Dim iRow As Integer
        Dim strFont As String
        Dim iSize  As Integer
   On Error GoTo ErrorManager
    grdServicios.Visible = False
    grdServicios.Redraw = False
        procolDatosProductoNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
        If Not procolDatosProductoNumero.MetConsultar(Me.proDatosProducto.proDetalleDatosProducto, False) Then
            MsgBox "No se pudieron consultar DatosProductoNumero"
            Exit Sub
        End If
        procolNovedadNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
        procolNovedadNumero.proIncidentId = proDatosProducto.proIncidentId
        If Not procolNovedadNumero.MetConsultar(proDatosProducto.proDetalleDatosProducto, False) Then
            MsgBox "No se pudieron consultar lNovedadNumero"
            Exit Sub
        End If
        With grdServicios
            TotalFilas = procolNovedadNumero.Count
            'Llenar filas de datosProducto
            For iContadorPro = 1 To procolDatosProductoNumero.Count
                bEsta = False
                For iContadorNov = 1 To procolNovedadNumero.Count
                    If procolDatosProductoNumero.Item(iContadorPro).proNumero = procolNovedadNumero.Item(iContadorNov).proNumero And _
                                                procolDatosProductoNumero.Item(iContadorPro).proRegionCode = procolNovedadNumero.Item(iContadorNov).proRegionCode Then
                        'ya está dentro de las novedades
                        bEsta = True
                    End If
                Next
                If Not bEsta Then
                        ' Llena las filas de DatosProductoNumero
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = 0
                        .TextMatrix(.Rows - 1, 1) = iCVienedeDatosProNum
                        .TextMatrix(.Rows - 1, 2) = procolDatosProductoNumero.Item(iContadorPro).proDatosProductoId
                        .TextMatrix(.Rows - 1, 3) = procolDatosProductoNumero.Item(iContadorPro).proRegionCode
                        .TextMatrix(.Rows - 1, 4) = procolDatosProductoNumero.Item(iContadorPro).proNumero
                End If
            Next
            'Llenar filas de NovedadNumero
            For iContadorNov = 1 To procolNovedadNumero.Count
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = 0
                    .TextMatrix(.Rows - 1, 1) = iCVienedeNovedadNumero
                    .TextMatrix(.Rows - 1, 2) = procolNovedadNumero.Item(iContadorNov).proNovedadNumeroId
                    .TextMatrix(.Rows - 1, 3) = procolNovedadNumero.Item(iContadorNov).proRegionCode
                    .TextMatrix(.Rows - 1, 4) = procolNovedadNumero.Item(iContadorNov).proNumero
            Next
            ' Actualiza los check de acuerdo a la base de datos
            For iRow = 2 To .Rows - 1
                For iCol = 5 To .Cols - 1
                    .Row = iRow
                    .Col = iCol
                    .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(iRow, iCol) = EstaSeleccionado(iRow, iCol, strFont, iSize)
                    .CellFontName = strFont
                    .CellFontSize = iSize
                Next iCol
            Next iRow
    End With
    grdServicios.Redraw = True
    grdServicios.Visible = True
      Exit Sub
ErrorManager:
    SubGMuestraError

End Sub

Private Function EstaSeleccionado(iRow As Integer, iCol As Integer, ByRef strFont As String, ByRef iSize As Integer) As String
   Dim sSiExiste As String
   Dim i As Integer
   
   
   On Error GoTo ErrorManager

    EstaSeleccionado = strUnChecked
    sSiExiste = "0"
    'Debemos saber por número que servicios suplementarios estan fijos y cuales estan en curso
    '1.buscamos en lo fijo
    procolServiciosxNumero.proDatosProductoId = proDatosProducto.proDatosProductoId
    procolServiciosxNumero.proServicioID = grdServicios.TextMatrix(0, iCol)
    procolServiciosxNumero.proRegionCode = grdServicios.TextMatrix(iRow, 3)
    procolServiciosxNumero.proNumero = grdServicios.TextMatrix(iRow, 4)
    If procolServiciosxNumero.MetConsultarxServicio Then
        sSiExiste = "1"
        'Si existen valores debe Buscar el Valor que tiene asignado para ese número
        'Buscar en la tabla CT_valorSERVICIOXNUMERO
        Me.procolValorServicioxNumero.proServicioSuplementarioId = procolServiciosxNumero.proServicioID
        Me.procolValorServicioxNumero.proNumero = procolServiciosxNumero.proNumero
        Me.procolValorServicioxNumero.proRegionCode = procolServiciosxNumero.proRegionCode
        If Me.procolValorServicioxNumero.MetConsultar Then
            If Me.procolValorServicioxNumero.Count > 0 Then
                If Me.procolValorServicioxNumero.Item(1).proTipoServicio = "C" Then 'Checkbox
                    EstaSeleccionado = strChecked
                    strFont = "Wingdings"
                    iSize = 14
                Else
                    If procolValorServicioxNumero.Count > 0 Then
                        EstaSeleccionado = procolValorServicioxNumero.Item(1).proValor
                        strFont = "Arial Narrow"
                        iSize = 8.5
                    Else
                        EstaSeleccionado = strUnChecked
                        strFont = "Wingdings"
                        iSize = 14
                    End If
                End If
            End If
        End If
    End If
    '2.buscamos en lo que va en curso
    For i = 1 To procolNovedadNumero.Count
        If grdServicios.TextMatrix(iRow, 4) = procolNovedadNumero.Item(i).proNumero And grdServicios.TextMatrix(iRow, 3) = procolNovedadNumero.Item(i).proRegionCode Then
            procolServiciosxReserva.proNovedadNumeroId = procolNovedadNumero.Item(i).proNovedadNumeroId
            procolServiciosxReserva.proServicioSuplementarioId = grdServicios.TextMatrix(0, iCol)
            procolServiciosxReserva.proRegionCode = grdServicios.TextMatrix(iRow, 3)
            procolServiciosxReserva.proNumero = grdServicios.TextMatrix(iRow, 4)
            If procolServiciosxReserva.MetConsultarxServicio Then
                sSiExiste = "1"
                'Si existen valores debe Buscar el Valor que tiene asignado para ese número
                'Buscar en la tabla CT_novedadvalorservicioxnumEro
                procolValorServicioxReserva.proNovedadNumeroId = procolServiciosxReserva.proNovedadNumeroId
                procolValorServicioxReserva.proServicioSuplementarioId = procolServiciosxReserva.proServicioSuplementarioId
                If procolValorServicioxReserva.MetConsultar Then
                    If Me.procolValorServicioxReserva.Count > 0 Then
                        If procolValorServicioxReserva.proTipoServicio = "C" Then 'Checkbox
                            EstaSeleccionado = strChecked
                            strFont = "Wingdings"
                            iSize = 14
                        Else
                            If procolValorServicioxReserva.Count > 0 Then
                                EstaSeleccionado = procolValorServicioxReserva.Item(1).proValor
                                strFont = "Arial Narrow"
                                iSize = 8.5
                            Else
                                EstaSeleccionado = strUnChecked
                                strFont = "Wingdings"
                                iSize = 14
                            End If
                        End If
                    End If
                Else
                    EstaSeleccionado = strUnChecked
                    strFont = "Wingdings"
                    iSize = 14
                End If
                Exit For
           
            End If
            
        End If
    Next
    If sSiExiste = "0" Then
        EstaSeleccionado = strUnChecked
        strFont = "Wingdings"
        iSize = 14
    End If
    
    Exit Function
ErrorManager:
    SubGMuestraError

End Function

Private Sub itmdessel_Click()
    Dim x As Integer
   On Error GoTo ErrorManager
    With grdServicios
        For x = 2 To .Rows - 1
        .Row = x
        If .TextMatrix(x, .Col) = strChecked Then
            .TextMatrix(x, .Col) = strUnChecked
            If .CellBackColor = strColorAzul Then
                .CellBackColor = vbWhite
            Else
                .CellBackColor = strColorAzul
            End If
        End If
        Next x
    End With
      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub itmSelall_Click()
    Dim x As Integer
   On Error GoTo ErrorManager
    With grdServicios
        For x = 2 To .Rows - 1
        .Row = x
        If .TextMatrix(x, .Col) = strUnChecked Then
            .TextMatrix(x, .Col) = strChecked
            If .CellBackColor = strColorAzul Then
                .CellBackColor = vbWhite
            Else
                .CellBackColor = strColorAzul
            End If
        End If
        Next x
    End With
      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub


