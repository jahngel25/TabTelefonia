Attribute VB_Name = "Function"
'Variable para saber si toca salir de la ventana de facturacion incident
'o toca cargarla en el momento de activar la ventana de facturacion
Dim varSalir As Boolean

Sub SubGMuestraErrorReservaDuplicado()
'***********************************************************
'   OBJETIVO:  Manejo de Errores Centralizado, captura el
'              error y lo despliega
'***********************************************************
'   AUTOR: Carlos Alberto Avila Gonzalez
'   FECHA: 14/11/2007
'***********************************************************
    Dim Descripcion As String
    Dim StartChar   As Integer
    
    StartChar = Len("[Microsoft][ODBC SQL Server Driver][SQL Server]")
    
    Descripcion = Mid(Err.Description, StartChar + 1, Len(Err.Description))
    
    MsgBox Descripcion, vbOKOnly + vbInformation, App.Title
End Sub



Sub SubGMuestraError()
'***********************************************************
'   OBJETIVO:  Manejo de Errores Centralizado, captura el
'              error y lo despliega
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 21/12/2000
'***********************************************************
    MsgBox "[" & Trim(Str(Err.Number)) & "] - " & _
        Err.Description, vbOKOnly + vbInformation, App.Title
End Sub
Function FunGFechaDMA(ParFecha As String) As String
'***********************************************************
'   OBJETIVO:  Toma una fecha en formato MM/DD/AAAA y la
'              convierte a DD/MM/AAAA
'************************************************************
'   PARAMETROS:  ParFecha       Fecvha en formato MM/DD/AAAA
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
Dim varFecha As String
Dim varResto As String
On Error GoTo ErrorManager


        If Trim(ParFecha) = "" Then Exit Function

        varFecha = Left(ParFecha, 10)
        varResto = Right(ParFecha, Len(ParFecha) - 10)
        
        FunGFechaDMA = Mid(varFecha, 4, 2) & "/" & _
                       Left(varFecha, 2) & "/" & _
                       Right(varFecha, 4) & varResto
                       
            
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
Function FunGFechaMDA(ParFecha As String) As String
'***********************************************************
'   OBJETIVO:  Toma una fecha en formato DD/MM/AAAA y la convierte a
'              MM/DD/AAAA
'************************************************************
'   PARAMETROS:  ParFecha       Fecvha en formato DD/MM/AAAA
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
Dim varFecha As String
Dim varResto As String
On Error GoTo ErrorManager

        
        If Trim(ParFecha) = "" Then Exit Function
        
        varFecha = Left(ParFecha, 10)
        varResto = Right(ParFecha, Len(ParFecha) - 10)
            
        FunGFechaMDA = Mid(varFecha, 4, 2) & "/" & _
                       Left(varFecha, 2) & "/" & _
                       Right(varFecha, 4) & varResto
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
Function FunGLeeNumerico(parAscii As Integer) As Integer
'***********************************************************
'   OBJETIVO:  Manejo de Errores Centralizado, captura el
'              error y lo despliega
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
On Error GoTo ErrorManager

    If (parAscii >= 48 And parAscii <= 57) Or parAscii = 8 Then
        FunGLeeNumerico = parAscii
    End If
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Function FunGLeeDireccion(parAscii As Integer, parControl As TextBox) As Integer
'***********************************************************
'   OBJETIVO:  Leer Caractéres Mayúsculas, Minúsculas, Guión,
'              No se permiten dobles, espacios
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
Dim varCaracter As String
On Error GoTo ErrorManager

    If (parAscii >= 48 And parAscii <= 57) Or _
        (parAscii >= 65 And parAscii <= 90) Or _
         (parAscii >= 97 And parAscii <= 122) Or _
                parAscii = 8 Or parAscii = 45 Or parAscii = 32 Then
        If parAscii = 32 Then 'Espacio a evaluar
            If Right(parControl.Text, 1) = " " Then Exit Function
        End If
        FunGLeeDireccion = parAscii
    End If
    
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Function FunGLeeAlfaNumerico(parAscii As Integer, Optional parCaseSensitive As Variant) As Integer
'***********************************************************
'   OBJETIVO:  Manejo de Errores Centralizado, captura el
'              error y lo despliega
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
Dim varCaracter As String
On Error GoTo ErrorManager

    'Convierte a Mayúsculas el caracter
    varCaracter = Chr(parAscii)
    If IsMissing(parCaseSensitive) = False Then
        If parCaseSensitive = False Then
            varCaracter = UCase(varCaracter)
        End If
    Else
        varCaracter = UCase(varCaracter)
    End If
    parAscii = Asc(varCaracter)
    If IsMissing(parCaseSensitive) = False Then
            If parCaseSensitive Then
                If (parAscii >= 97 And parAscii <= 122) Or (parAscii >= 40 And parAscii <= 57) Or (parAscii >= 65 And parAscii <= 90) Or _
                            parAscii = 241 Or parAscii = 8 Or parAscii = 32 Or parAscii = 35 Or parAscii = 38 Or parAscii = 95 Then
                    FunGLeeAlfaNumerico = parAscii
                End If
            Else
                If (parAscii >= 40 And parAscii <= 57) Or (parAscii >= 65 And parAscii <= 90) Or _
                            parAscii = 209 Or parAscii = 8 Or parAscii = 32 Or parAscii = 35 Or parAscii = 38 Or parAscii = 95 Then
                    FunGLeeAlfaNumerico = parAscii
                End If
            End If
    Else
        If (parAscii >= 40 And parAscii <= 57) Or (parAscii >= 65 And parAscii <= 90) Or _
                            parAscii = 209 Or parAscii = 8 Or parAscii = 32 Or parAscii = 35 Or parAscii = 38 Or parAscii = 95 Then
                    FunGLeeAlfaNumerico = parAscii
                End If
    End If
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Function FunGVerificaDatos(ParamArray parControles() As Variant) As Boolean
'**************************'***********************************************************
'   OBJETIVO:  Encuentra si falta algún dato en los controles,
'              en caso contrario muestra un mensaje y retorna el
'              foco al control
'**************************'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 29/12/2000
'**************************'***********************************************************
Dim varContador As Integer
Dim varEncontro As Integer

On Error GoTo ErrorManager

    varContador = 0
    varEncontro = False
    While varContador <= UBound(parControles) And varEncontro = False
        If TypeOf parControles(varContador) Is TextBox Then
            If Len(Trim(parControles(varContador))) = 0 Then varEncontro = True
        End If
        If TypeOf parControles(varContador) Is ComboBox Then
            If parControles(varContador).ListIndex = -1 Then varEncontro = True
        End If
        If TypeOf parControles(varContador) Is DTPicker Then
            If IsNull(parControles(varContador).Value) = True Then varEncontro = True
        End If
        
        If varEncontro = False Then
            varContador = varContador + 1
        Else
            MsgBox parControles(varContador).Tag, vbOKOnly + vbInformation, App.Title
            parControles(varContador).SetFocus
            Exit Function
        End If
    Wend
    
    FunGVerificaDatos = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Function FunGValidarProducto(parProducto As Long, _
                              parConexion As ADODB.Connection) As String
'**************************'***********************************************************
'   OBJETIVO:  Valida si el producto pertence al cliente
'**************************'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 29/12/2000
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    varComando = "select vchSerialNumber, " & _
                 "       iOwnerId         " & _
                 "from   Customer_Product " & _
                 "where  iProductId = " & parProducto
            
    varResultado.Open varComando, parConexion
    
    If varResultado.EOF Then
        MsgBox "El producto Digitado no existe. No se almacenara la información del Asunto.", vbCritical, App.Title
        FunGValidarProducto = ""
    Else
        If IsNull(varResultado.Fields("iOwnerId")) Then
            MsgBox "El Producto No tiene Cliente Asociado. No se almacenara la información del Asunto.", vbInformation, App.Title
            FunGValidarProducto = ""
        Else
            If varResultado.Fields("iOwnerId") <> CLng(Trim(frmFacturacion.proCompanyId)) Then
                MsgBox "El producto no pertece a este Cliente. No se almacenara la información del Asunto.", vbInformation, App.Title
                FunGValidarProducto = ""
            Else
                If IsNull(varResultado.Fields("vchSerialNumber")) Then
                    MsgBox "El producto no tiene codigo de enlace asociado. No se almacenara la información del Asunto.", vbInformation, App.Title
                    FunGValidarProducto = ""
                Else
                    FunGValidarProducto = varResultado.Fields("vchSerialNumber")
                End If
            End If
        End If
    End If
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Function FunGValidarExistenciaAsunto(parAsunto As String, _
                                                            parConexion As ADODB.Connection) As Boolean
                                                            
'**************************'***********************************************************
'   OBJETIVO:  Valida si el asunto digitado ya fue utilizado en otra facturacion
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    'Verificar si el incidente ya esta ligado a alguna facturacion
    varComando = "select count(*) " & _
                         "from   ct_FacturacionIncident " & _
                         "where  iIncidentId = " & parAsunto
            
    varResultado.Open varComando, parConexion
    
    If varResultado.EOF Then
        FunGValidarExistenciaAsunto = True
    Else
        If IsNull(varResultado.Fields(0)) Then
             FunGValidarExistenciaAsunto = True
        Else
            If varResultado.Fields(0) <> 0 Then
                MsgBox "El asunto digitado ya fue utilizado en alguna facturación. ", vbInformation, App.Title
                FunGValidarExistenciaAsunto = False
            Else
                FunGValidarExistenciaAsunto = True
            End If
        End If
    End If
    
    Set varResultado = Nothing
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Function FunGValidarClienteAsunto(parAsunto As String, _
                                                        parConexion As ADODB.Connection) As Boolean
        
'**************************'***********************************************************
'   OBJETIVO:  Valida si el asunto pertenece al cliente actual
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
        
    'Verificar si el asunto pertenece al cliente actual y cual es la categoria del incidente
    varComando = "select   iOwnerId            " & _
                          "from    Incident                " & _
                          "where  iIncidentId = " & parAsunto
            
    varResultado.Open varComando, parConexion
    
    If varResultado.EOF Then
        MsgBox "El Asunto Digitado no existe. ", vbCritical, App.Title
        FunGValidarClienteAsunto = False
    Else
        If IsNull(varResultado.Fields("iOwnerId")) Then
            MsgBox "El Asunto No tiene Cliente Asociado. ", vbCritical, App.Title
            FunGValidarClienteAsunto = False
        Else
            If varResultado.Fields("iOwnerId") <> CLng(Trim(frmVoz.proCompanyId)) Then
                MsgBox "El Asunto no pertece a este Cliente. ", vbCritical, App.Title
                FunGValidarClienteAsunto = False
            Else
                 FunGValidarClienteAsunto = True
            End If
        End If
    End If
    Set varResultado = Nothing
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Function FunGValidarCategoriaAsunto(parAsunto As String, _
                                                            parConexion As ADODB.Connection) As Boolean
                                                            
'**************************'***********************************************************
'   OBJETIVO:  Validar la Categoria del asunto digitado
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim Producto As Double
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
        
    varComando = "select   iIncidentCategory            " & _
                          "from    Incident                          " & _
                          "where  iIncidentId = " & parAsunto
            
    varResultado.Open varComando, parConexion
    
    If varResultado.EOF Then
        MsgBox "El Asunto Digitado no existe. ", vbCritical, App.Title
        FunGValidarCategoriaAsunto = False
        Exit Function
    Else
        If IsNull(varResultado.Fields("iIncidentCategory")) Then
            MsgBox "El Asunto No tiene Categoria Asociada. ", vbCritical, App.Title
            FunGValidarCategoriaAsunto = False
            Exit Function
        Else
            If varResultado.Fields("iIncidentCategory") = 3 Then
                If Val(frmFacturacion.proFacturacion.proFacturacionId) <> 0 Then
                    MsgBox "Esta facturacion ya tiene ligado un Lead. Debe crear una " & Chr(13) & _
                                 "nueva facturación o seleccionar el ticket que modifique esta.", vbInformation, App.Title
                    FunGValidarCategoriaAsunto = False
                    Exit Function
                End If
            End If
                    
            If Val(varResultado.Fields("iIncidentCategory")) <> 2 And Val(varResultado.Fields("iIncidentCategory")) <> 3 Then
                MsgBox "La facturacion solo puede estar ligada a incidentes de ventas o de atencion.", vbInformation, App.Title
                FunGValidarCategoriaAsunto = False
                Exit Function
            End If
                    
            If Val(varResultado.Fields("iIncidentCategory")) = 2 Then
                If Val(Trim(frmFacturacion.proFacturacion.proFacturacionId)) = 0 Then
                    MsgBox "Solo se puede generar una facturacion nueva por un incidente de ventas.", vbInformation, App.Title
                    FunGValidarCategoriaAsunto = False
                    Exit Function
                Else
                    Producto = FunGBuscarProductoAsunto(parAsunto, parConexion)
                    If Producto <> Trim(frmFacturacion.proFacturacion.proProductId) Then
                        MsgBox "El incidente de atencion digitado no pertenece al servicio de esta facturacion.", vbInformation, App.Title
                        FunGValidarCategoriaAsunto = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    Set varResultado = Nothing
    FunGValidarCategoriaAsunto = True
    Exit Function
ErrManager:
    FunGValidarCategoriaAsunto = False
    SubGMuestraError
End Function
Function FunGBuscarProductoAsunto(parAsunto As String, _
                                                            parConexion As ADODB.Connection) As String

'**************************'***********************************************************
'   OBJETIVO:  Busca el producto ligado para un asunto de atencion
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
        
    varComando = "Select CP.iProductId    " & _
                            "From   Customer_Product CP, " & _
                            "           Incident I                    " & _
                            "Where I.iIncidentId = " & Val(parAsunto) & " " & _
                            "And     I.vchUser9 = CP.vchSerialNumber "
                            
    varResultado.Open varComando, parConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("iProductId").Value) Then
            FunGBuscarProductoAsunto = ""
        Else
            FunGBuscarProductoAsunto = Trim(varResultado.Fields("iProductId").Value)
        End If
    End If
    
    Set varResultado = Nothing
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Function FunGValidarOTCerrada(parAsunto As String, _
                                                   parConexion As ADODB.Connection) As Boolean
                                                            
'**************************'***********************************************************
'   OBJETIVO:  Valida si el asunto digitado tiene la ot de instalacion cerrada o no
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    'Verificar si el incidente ya esta ligado a alguna facturacion
    varComando = "select iStatusId " & _
                         "from    Incident " & _
                         "where  vchUser7 = " & Val(Trim(parAsunto)) & _
                         " and iIncidentTypeId = 102414 " & _
                         " and tiRecordStatus = 1"
            
    varResultado.Open varComando, parConexion
    
    If varResultado.EOF Then
        FunGValidarOTCerrada = True
    Else
        If IsNull(varResultado.Fields(0)) Or varResultado.Fields(0) = 0 Then
             MsgBox "La Ot no tiene estado definido.", vbCritical, App.Title
             FunGValidarOTCerrada = False
        Else
            Select Case varResultado.Fields(0)
                Case 101456
                    MsgBox "La OT del asunto se encuentra suspendida.", vbInformation, App.Title
                    FunGValidarOTCerrada = False
                Case 101465
                    MsgBox "La OT del asunto se encuentra cerrada.", vbInformation, App.Title
                    FunGValidarOTCerrada = False
                Case 103679
                    MsgBox "La OT del asunto se encuentra cancelada.", vbInformation, App.Title
                    FunGValidarOTCerrada = False
                Case Else
                    FunGValidarOTCerrada = True
            End Select
        End If
    End If
    
    Set varResultado = Nothing
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Function FunGValidarAsuntoTelefonia(parAsunto As String, _
                                     parConexion As ADODB.Connection) As Boolean
'**************************'***********************************************************
'   OBJETIVO:  valida que el asunto digitado no exista  sea
'                      un asunto de ventas  que este activo
'**************************'***********************************************************
'   AUTOR: Hernan Botache
'   FECHA: 02/12/2004
'**************************'***********************************************************
  On Error GoTo ErrManager
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
     
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
        
    varComando = "select    iIncidentCategory, iStatusId, tiRecordStatus  " & _
                          " from    Incident " & _
                          " where  iIncidentId = '" & parAsunto & "' and " & _
                          " vchUser5 in (select  iParameterId from referenceDefinition " & _
                          " where   vchExtradata = '1810' or vchExtradata = '1812')"
            
    varResultado.Open varComando, parConexion
    
    If varResultado.EOF = False Then
        'verifica el estado
        Select Case varResultado.Fields("iStatusId")
                  Case 101456
                      MsgBox "La OT del asunto se encuentra suspendida.", vbInformation, App.Title
                      FunGValidarAsuntoTelefonia = False
                  Case 101465
                      MsgBox "La OT del asunto se encuentra cerrada.", vbInformation, App.Title
                      FunGValidarAsuntoTelefonia = False
                  Case 103679
                      MsgBox "La OT del asunto se encuentra cancelada.", vbInformation, App.Title
                      FunGValidarAsuntoTelefonia = False
                  Case Else
                      FunGValidarAsuntoTelefonia = True
        End Select
        'verifica la categoria
        If IsNull(varResultado.Fields("iIncidentCategory")) Then
            MsgBox "El Asunto No tiene Categoria Asociada. ", vbCritical, App.Title
            FunGValidarAsuntoTelefonia = False
            Exit Function
        Else
            If varResultado.Fields("iIncidentCategory") <> 3 Then
                    MsgBox "La telefonia local solo puede estar ligada a incidentes de ventas", vbInformation, App.Title
                    FunGValidarAsuntoTelefonia = False
                    Exit Function
            End If
        End If
        'verifica el estado
        If varResultado.Fields("tiRecordStatus") = 0 Then
                MsgBox "La venta se encuentra deshabilitado", vbInformation, App.Title
                FunGValidarAsuntoTelefonia = False
                Exit Function
        End If
        If FunGVerificarBolsaCoorporativa(parAsunto, parConexion) = False Then
            If FunGVerificarLocalCoorporativa(parAsunto, parConexion) = False Then
                MsgBox "La venta local coorporativa se encuentra utilizada en otro producto de telefonia", vbInformation, App.Title
                FunGValidarAsuntoTelefonia = False
                Exit Function
            
            End If
        End If
    Else
        MsgBox "El asunto no existe, o no es un asunto de telefonia local", vbInformation, App.Title
        FunGValidarAsuntoTelefonia = False
        Exit Function
    End If
       
    FunGValidarAsuntoTelefonia = True
    Exit Function
ErrManager:
    SubGMuestraError

End Function
Function FunGVerificarLocalCoorporativa(parAsunto As String, _
                                                   parConexion As ADODB.Connection) As Boolean
 '**************************'***********************************************************
'   OBJETIVO:  Valida que no halla sido utlizada en otro producto
'**************************'***********************************************************
'   AUTOR: Hernan Botache
'   FECHA: 02/12/2004
'**************************'***********************************************************
  Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    'Verificar si el incidente ya esta ligado a alguna facturacion
    varComando = "select    iDatosProductoId  " & _
                          " from    CT_DatosProducto " & _
                          " where  iVentaId = '" & parAsunto & "'"
            
    varResultado.Open varComando, parConexion
    If varResultado.EOF = False Then
        FunGVerificarLocalCoorporativa = False
    Else
        FunGVerificarLocalCoorporativa = True
    End If
    
    Set varResultado = Nothing
    Exit Function
ErrManager:
    FunGVerificarLocalCoorporativa = False
    SubGMuestraError
End Function
Function FunGVerificarBolsaCoorporativa(parAsunto As String, _
                                                   parConexion As ADODB.Connection) As Boolean
                                                            
'**************************'***********************************************************
'   OBJETIVO:  Valida si el asunto es de bolsa coorporativa
'**************************'***********************************************************
'   AUTOR: Hernan Botache
'   FECHA: 02/12/2004
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    'Verificar si el incidente ya esta ligado a alguna facturacion
    varComando = "select    iIncidentCategory, iStatusId, tiRecordStatus  " & _
                          " from    Incident " & _
                          " where  iIncidentId = '" & parAsunto & "' and " & _
                          " vchUser5 in (select  iParameterId from referenceDefinition " & _
                          " where    vchExtradata = '1812')"
            
    varResultado.Open varComando, parConexion
    
    
    
    If varResultado.EOF = False Then
        FunGVerificarBolsaCoorporativa = True
    Else
        FunGVerificarBolsaCoorporativa = False
    End If
    
    Set varResultado = Nothing
    Exit Function
ErrManager:
    FunGVerificarBolsaCoorporativa = False
    SubGMuestraError
End Function

Function FunGValidarAsunto(parAsunto As String, _
                           parConexion As ADODB.Connection) As Boolean
'**************************'***********************************************************
'   OBJETIVO:  valida que el asunto digitado no exista en otra facturacion, sea
'                      un asunto de ventas o de atencion, que pertenezca al usuario
'                      activo
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************

    On Error GoTo ErrManager
    
    If FunGValidarExistenciaAsunto(parAsunto, parConexion) Then
        If FunGValidarClienteAsunto(parAsunto, parConexion) Then
            If FunGValidarCategoriaAsunto(parAsunto, parConexion) Then
                If Not FunGValidarOTCerrada(parAsunto, parConexion) Then
                    MsgBox "No se puede editar la información del Asunto.", vbInformation, App.Title
                    FunGValidarAsunto = False
                    Exit Function
                End If
            Else
                MsgBox "No se puede editar la información del Asunto.", vbInformation, App.Title
                FunGValidarAsunto = False
                Exit Function
            End If
        Else
            MsgBox "No se puede editar la información del Asunto.", vbInformation, App.Title
            FunGValidarAsunto = False
            Exit Function
        End If
    Else
        MsgBox "No se puede editar la información del Asunto.", vbInformation, App.Title
        FunGValidarAsunto = False
        Exit Function
    End If
    
    FunGValidarAsunto = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Sub SubFPintarFila(grdGrid As MSFlexGrid, parFila As Integer, Optional parPintar As Variant)
    Dim Columna As Integer
    On Error GoTo ErrManager
    'Cambiar Color Fila
    If Not IsMissing(parPintar) Then
        For Columna = 0 To grdGrid.Cols - 1
            grdGrid.Row = parFila
            grdGrid.Col = Columna
            grdGrid.CellBackColor = parPintar
        Next Columna
    Else
        'Reestablecer color fila
        For Columna = 0 To grdGrid.Cols - 1
            grdGrid.Row = parFila
            grdGrid.Col = Columna
            grdGrid.CellBackColor = parPintar
        Next Columna
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Function FunGQuitarTab(ParCadena As String) As String
'***********************************************************
'   OBJETIVO:  Quitar los tabs que aparescan en una cadena
'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 02/04/2001
'***********************************************************
    Dim Contador As Integer
    Dim Longitud As Integer
    Dim Resultado As String
    Dim varLetra As String
    On Error GoTo ErrManager
    
    Contador = 1
    Longitud = Len(ParCadena)
        
    While Contador <= Longitud
        varLetra = Mid(ParCadena, Contador, 1)
        If varLetra = Chr(9) Then
            varResultado = varResultado & " "
        Else
            varResultado = varResultado & varLetra
        End If
        Contador = Contador + 1
    Wend
    FunGQuitarTab = varResultado
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Function FunGLeeDecimales(KeyAscii As Integer, parTexto As TextBox) As Integer
'************************************************************************
'*  OBJETIVOS :  Lee la cantidad indicada de enteros y decimales
'************************************************************************
'*  PARAMETROS:
'*      KeyAscii                Ascii de la tecla oprimida
'*      ParTexto                Texto a validar
'*
'*  RESULTADOS:
'*      #                           Tecla Válida
'*      0                           Tecla Inválida
'*************************************************************************
'*  SONDA de Colombia
'*  Autor: Raúl Cruz A.
'*  Fecha: 06 / 02 / 2001
'***********************************************************************
Dim varEnteros As Integer
Dim varDecimales As Integer
Dim varEnterosActuales As Integer
Dim varDecimalesActuales As Integer
Dim varValorMaximo As Double
Dim varPosicionPunto As Integer
Dim varValor As Double
Dim varTexto As String
Dim varTextoSeleccion As String
On Error GoTo ErrorManager

    'valida que los caracteres sean numéricos
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Then
       FunGLeeDecimales = 0
    Else
       Exit Function
    End If
    
    parTexto = Trim(parTexto)
    varTextoSeleccion = Left(parTexto, parTexto.SelStart) + Right(parTexto, Len(parTexto) - (parTexto.SelStart + parTexto.SelLength))
    
    If Len(Trim(parTexto.Tag)) > 0 Then
        'Toma el número de enteros y Decimales
        varPosicionPunto = InStr(parTexto.Tag, ".")
        If varPosicionPunto = 0 Then
            varPosicionPunto = InStr(parTexto.Tag, ",")
        End If
        If varPosicionPunto Then
            varDecimales = Len(Trim(parTexto.Tag)) - varPosicionPunto
            varEnteros = varPosicionPunto - 1
        Else
            varEnteros = Len(Trim(parTexto.Tag))
        End If
        
        'Caso en que el ascii es el del punto o la coma
        If Chr(KeyAscii) = "." Or Chr(KeyAscii) = "," Then
                'Averigua si debe recibir coma o punto
                If varDecimales = 0 Then Exit Function
                
                'Averigua si ya existia una coma o punto
                If InStr(varTextoSeleccion, ".") > 0 Or InStr(varTextoSeleccion, ",") Then Exit Function
                
                If Len(Trim(varTextoSeleccion)) = 0 Then
                        parTexto = "0"
                        parTexto.SelStart = 1
                End If
        End If
        
        'Busca el valor máximo a leer
        varValorMaximo = Val(Left(parTexto.Tag, varEnteros))
        If varDecimales Then
            varValorMaximo = varValorMaximo + Val("0." & Right(parTexto.Tag, varDecimales))
        End If
        
        'Construye la cadena con el valor si fuera incluido
        varTexto = Left(parTexto, parTexto.SelStart)  'Izquierda
        If KeyAscii <> 8 Then
                varTexto = varTexto & Chr(KeyAscii)                 'Nuevo caracter
        Else
                If parTexto.SelLength = 0 Then
                        varTexto = Left(varTexto, Len(varTexto) - 1)
                End If
        End If
        varTexto = varTexto & Right(parTexto, Len(Trim(parTexto)) - (parTexto.SelStart + parTexto.SelLength))  'Derecha
        
        'Busca el valor de la cadena incluyendo el caracter
        varPosicionPunto = InStr(varTexto, ".")
        If varPosicionPunto = 0 Then
            varPosicionPunto = InStr(varTexto, ",")
        End If
        If varPosicionPunto Then
            varDecimalesActuales = Len(Trim(varTexto)) - varPosicionPunto
            varEnterosActuales = varPosicionPunto - 1
        Else
            varEnterosActuales = Len(Trim(varTexto))
        End If
        
        'No puede tener más decimales de los estipulados, ni enteros de los estipulados
        If varDecimalesActuales > varDecimales Then Exit Function
        If varEnterosActuales > varEnteros Then Exit Function
        
        'Construye el valor
        varValor = Val(Left(varTexto, varEnterosActuales))
        If varDecimalesActuales Then
            varValor = varValor + Val("0." & Right(varTexto, varDecimalesActuales))
        End If
        
        'Si el valor actual + el nuevo caracter supera al máximo,
        'Retorna 0
        If varValor > varValorMaximo Then Exit Function
    End If
    
    'Retorna el valor del KeyAscii
    FunGLeeDecimales = KeyAscii
    Exit Function

ErrorManager:
    SubGMuestraError
End Function

Function FunGFechaAMD(ParFecha As String) As String
'***********************************************************
'   OBJETIVO:  Toma una fecha en formato DD/MM/AAAA y la
'              convierte a AAAA/MM/DD
'************************************************************
'   PARAMETROS:  ParFecha       Fecha en formato DD/MM/AAAA
'***********************************************************
'   AUTOR: TOPGROUP S.A.
'   FECHA: 05/08/2009
'   VERSION: 1.0.100
'   REQUERIMIENTO:3488
'***********************************************************
Dim varFecha As Variant
Dim varResto As Variant
On Error GoTo ErrorManager


        If Trim(ParFecha) = "" Then Exit Function

        varFecha = Left(ParFecha, 10)
        varResto = Right(ParFecha, Len(ParFecha) - 10)
        
        FunGFechaAMD = Right(varFecha, 4) & "/" & _
                       Mid(varFecha, 4, 2) & "/" & _
                       Left(varFecha, 2) & varResto
                       
            
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

