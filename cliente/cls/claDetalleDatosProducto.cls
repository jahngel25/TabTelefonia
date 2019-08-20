VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claDetalleDatosProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proDatosProductoId As String
Public proDetalleDatosProductoId As String
Public proStatusId As String
Public proUser1 As String
Public proUser2 As String
Public proUser3 As String
Public proUser4 As String
Public proUser5 As String
Public proUser6 As String
Public proUser7 As String
Public proUser8 As String
Public proUser9 As String
Public proUser10 As String
Public proUser11 As String
Public proUser12 As String
Public proUser13 As String
Public proUser14 As String
Public proUser15 As String
Public proUser16 As String
Public proUser17 As String
Public proUser18 As String
Public proUser19 As String
Public proUser20 As String
Public proUser21 As String
Public proUser22 As String
Public proUser23 As String
Public proUser24 As String
Public proUser25 As String
Public proUser26 As String
Public proUser27 As String
Public proUser28 As String
Public proUser29 As String
Public proUser30 As String
Public proUser31 As String
Public proUser32 As String
Public proUser33 As String
Public proUser34 As String
Public proUser35 As String
Public proUser36 As String
Public proUser37 As String
Public proUser38 As String
Public proUser39 As String
Public proUser40 As String

Public proRecordStatus As String
Public proSeleccion As String

Public proParametrosxProducto As EDCAdminVoz.colParametroProducto
''Modificaciones para relacion de numeros con tipos de linea
Public proContador As Integer
Public proEliminar As Boolean
Public proBackUp As Boolean
Public proModificar As Boolean
Public proConexion As ADODB.Connection

Public Function FunGAumentarContador() As Boolean
    proContador = proContador + 1
End Function
Public Function FunGDisminuirContador() As Boolean
    proContador = proContador - 1
End Function

Public Function MetInsertar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    varComando = "INSERT INTO CT_DetalleDatosProducto ( iDatosProductoId,   chStatusId, " & _
                 "            vchUser1,     vchUser2,   vchUser3,   vchUser4,           " & _
                 "            vchUser5,     vchUser6,   vchUser7,   vchUser8,           " & _
                 "            vchUser9,     vchUser10,  vchUser11,  vchUser12,          " & _
                 "            vchUser13,    vchUser14,  vchUser15,  vchUser16,          " & _
                 "            vchUser17,    vchUser18,  vchUser19,  vchUser20,          " & _
                 "            vchUser21,    vchUser22,  vchUser23,  vchUser24,          " & _
                 "            vchUser25,    vchUser26,  vchUser27,  vchUser28,          " & _
                 "            vchUser29,    vchUser30,  vchUser31,  vchUser32,          " & _
                 "            vchUser33,    vchUser34,  vchUser35,  vchUser36,          " & _
                 "            vchUser37,    vchUser38,  vchUser39,  vchUser40,          " & _
                 "            tiRecordStatus )                                          " & _
                 "VALUES (" & Me.proDatosProductoId & ", 'A" & "', '" & _
                              Me.proUser1 & "', '" & Me.proUser2 & "', '" & Me.proUser3 & "', '" & Me.proUser4 & "', '" & _
                              Me.proUser5 & "', '" & Me.proUser6 & "', '" & Me.proUser7 & "', '" & Me.proUser8 & "', '" & _
                              Me.proUser9 & "', '" & Me.proUser10 & "', '" & Me.proUser11 & "', '" & Me.proUser12 & "', '" & _
                              Me.proUser13 & "', '" & Me.proUser14 & "', '" & Me.proUser15 & "', '" & Me.proUser16 & "', '" & _
                              Me.proUser17 & "', '" & Me.proUser18 & "', '" & Me.proUser19 & "', '" & Me.proUser20 & "', " & _
                              Me.proUser21 & "', '" & Me.proUser22 & "', '" & Me.proUser23 & "', '" & Me.proUser24 & "', " & _
                              Me.proUser25 & "', '" & Me.proUser26 & "', '" & Me.proUser27 & "', '" & Me.proUser28 & "', " & _
                              Me.proUser29 & "', '" & Me.proUser30 & "', '" & Me.proUser31 & "', '" & Me.proUser32 & "', " & _
                              Me.proUser33 & "', '" & Me.proUser34 & "', '" & Me.proUser35 & "', '" & Me.proUser36 & "', " & _
                              Me.proUser37 & "', '" & Me.proUser38 & "', '" & Me.proUser39 & "', '" & Me.proUser40 & "', " & _
                              Me.proRecordStatus & ")"
                 
    Set varResultado = New ADODB.Recordset
    
    Me.proConexion.Execute varComando
    
    varComando = "Select Max(iDetalleDatosProductoId) " & _
                 "From   CT_DetalleDatosProducto "
                 
    varResultado.Open varComando, Me.proConexion
    
    If IsNull(varResultado.Fields(0)) Then
        Me.proDetalleDatosProductoId = 0
    Else
        Me.proDetalleDatosProductoId = Trim(varResultado.Fields(0))
    End If
    
    Set varResultado = Nothing
    MetInsertar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetActualizar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    Me.proStatusId = "A"
    varComando = "UPDATE    CT_DetalleDatosProducto " & _
                 "SET chStatusId = '" & Me.proStatusId & "', " & _
                 "    vchUser1   = '" & Me.proUser1 & "', " & "  vchUser2  = '" & Me.proUser2 & "', " & "  vchUser3   = '" & Me.proUser3 & "', " & "  vchUser4  = '" & Me.proUser4 & "', " & _
                 "    vchUser5   = '" & Me.proUser5 & "', " & "  vchUser6  = '" & Me.proUser6 & "', " & "  vchUser7   = '" & Me.proUser7 & "', " & "  vchUser8  = '" & Me.proUser8 & "', " & _
                 "    vchUser9   = '" & Me.proUser9 & "', " & "  vchUser10 = '" & Me.proUser10 & "', " & " vchUser11  = '" & Me.proUser11 & "', " & " vchUser12 = '" & Me.proUser12 & "', " & _
                 "    vchUser13  = '" & Me.proUser13 & "', " & " vchUser14 = '" & Me.proUser14 & "', " & " vchUser15  = '" & Me.proUser15 & "', " & " vchUser16 = '" & Me.proUser16 & "', " & _
                 "    vchUser17  = '" & Me.proUser17 & "', " & " vchUser18 = '" & Me.proUser18 & "', " & " vchUser19  = '" & Me.proUser19 & "', " & " vchUser20 = '" & Me.proUser20 & "', " & _
                 "    vchUser21  = '" & Me.proUser21 & "', " & " vchUser22 = '" & Me.proUser22 & "', " & " vchUser23  = '" & Me.proUser23 & "', " & " vchUser24 = '" & Me.proUser24 & "', " & _
                 "    vchUser25  = '" & Me.proUser25 & "', " & " vchUser26 = '" & Me.proUser26 & "', " & " vchUser27  = '" & Me.proUser27 & "', " & " vchUser28 = '" & Me.proUser28 & "', " & _
                 "    vchUser29  = '" & Me.proUser29 & "', " & " vchUser30 = '" & Me.proUser30 & "', " & " vchUser31  = '" & Me.proUser31 & "', " & " vchUser32 = '" & Me.proUser32 & "', " & _
                 "    vchUser33  = '" & Me.proUser33 & "', " & " vchUser34 = '" & Me.proUser34 & "', " & " vchUser35  = '" & Me.proUser35 & "', " & " vchUser36 = '" & Me.proUser36 & "', " & _
                 "    vchUser37  = '" & Me.proUser37 & "', " & " vchUser38 = '" & Me.proUser38 & "', " & " vchUser39  = '" & Me.proUser39 & "', " & " vchUser40 = '" & Me.proUser40 & "', " & _
                 "    tiRecordStatus  =  " & Me.proRecordStatus & " " & _
                 "WHERE     iDatosProductoId = " & Me.proDatosProductoId & " " & _
                 "AND       iDetalleDatosProductoId = " & Me.proDetalleDatosProductoId
        
    Me.proConexion.Execute varComando
    
    MetActualizar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "DELETE FROM CT_DetalleDatosProducto " & _
                 "WHERE iDatosProductoId = " & Me.proDatosProductoId & " " & _
                 "AND   iDetallesDatosProductoId = " & Me.proDetalleDatosProductoId
                 
    Me.proConexion.Execute varComando
    
    MetEliminar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetGuardar() As Boolean
    On Error GoTo ErrManager
        
        If Val(Trim(Me.proDetalleDatosProductoId)) = 0 Then
            If Me.MetInsertar Then
                MetGuardar = True
            Else
                MetGuardar = False
            End If
        Else
            If Me.MetActualizar Then
                MetGuardar = True
            Else
                MetGuardar = False
            End If
        End If
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetValidarValor(ByRef parCampo As String) As Boolean
'Metodo que valida si los valores registrados en el detalle son validos
'o si se requiere cambiar alguno, porque ya fue usado en algun otro
'servicio. El metodo retorna verdadero si se puede almacenar el registro
'y falso si no. El parametro parCampo es de salida y retorna el campo con
'el valor invalido

    Dim varContador As Integer
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    Set varComando = New ADODB.Command
    Set varResultado = New ADODB.Recordset
    
    Set varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CP_ValidarExistenciaValorCampoProducto"
    varComando.CommandType = adCmdStoredProc
    varComando.Parameters.Item("@iDatosProductoId").Value = Me.proDatosProductoId
    
    MetValidarValor = True
    
    'Recorrer la coleccion de parametros verificando cuales campos se deben validar
    For varContador = 1 To Me.proParametrosxProducto.Count
        If Me.proParametrosxProducto.Item(varContador).proValidarRepetidos = "1" Then
            
            varComando.Parameters("@chCampo").Value = Trim(Me.proParametrosxProducto.Item(varContador).proCampo)
            varComando.Parameters.Item("@chProductNumber").Value = Me.proParametrosxProducto.Item(varContador).proProductNumber
            
            Select Case Trim(Me.proParametrosxProducto.Item(varContador).proCampo)
                Case "vchUser1"
                    If Trim(Me.proUser1) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser1
                    End If
                Case "vchUser2"
                    If Trim(Me.proUser2) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser2
                    End If
                Case "vchUser3"
                    If Trim(Me.proUser3) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser3
                    End If
                Case "vchUser4"
                    If Trim(Me.proUser4) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser4
                    End If
                Case "vchUser5"
                    If Trim(Me.proUser5) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser5
                    End If
                Case "vchUser6"
                    If Trim(Me.proUser6) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser6
                    End If
                Case "vchUser7"
                    If Trim(Me.proUser7) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser7
                    End If
                Case "vchUser8"
                    If Trim(Me.proUser8) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser8
                    End If
                Case "vchUser9"
                    If Trim(Me.proUser9) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser9
                    End If
                Case "vchUser10"
                    If Trim(Me.proUser10) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser10
                    End If
                Case "vchUser11"
                    If Trim(Me.proUser11) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser11
                    End If
                Case "vchUser12"
                    If Trim(Me.proUser12) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser12
                    End If
                Case "vchUser13"
                    If Trim(Me.proUser13) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser13
                    End If
                Case "vchUser14"
                    If Trim(Me.proUser14) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser14
                    End If
                Case "vchUser15"
                    If Trim(Me.proUser15) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser15
                    End If
                Case "vchUser16"
                    If Trim(Me.proUser16) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser16
                    End If
                Case "vchUser17"
                    If Trim(Me.proUser17) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser17
                    End If
                Case "vchUser18"
                    If Trim(Me.proUser18) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser18
                    End If
                Case "vchUser19"
                    If Trim(Me.proUser19) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser19
                    End If
                Case "vchUser20"
                    If Trim(Me.proUser20) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser20
                    End If
                Case "vchUser21"
                    If Trim(Me.proUser21) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser21
                    End If
                Case "vchUser22"
                    If Trim(Me.proUser22) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser22
                    End If
                Case "vchUser23"
                    If Trim(Me.proUser23) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser23
                    End If
                Case "vchUser24"
                    If Trim(Me.proUser24) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser24
                    End If
                Case "vchUser25"
                    If Trim(Me.proUser25) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser25
                    End If
                Case "vchUser26"
                    If Trim(Me.proUser26) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser26
                    End If
                Case "vchUser27"
                    If Trim(Me.proUser27) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser27
                    End If
                Case "vchUser28"
                    If Trim(Me.proUser28) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser28
                    End If
                Case "vchUser29"
                    If Trim(Me.proUser29) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser29
                    End If
                Case "vchUser30"
                    If Trim(Me.proUser30) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser30
                    End If
                Case "vchUser31"
                    If Trim(Me.proUser31) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser31
                    End If
                Case "vchUser32"
                    If Trim(Me.proUser32) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser32
                    End If
                Case "vchUser33"
                    If Trim(Me.proUser33) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser33
                    End If
                Case "vchUser34"
                    If Trim(Me.proUser34) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser34
                    End If
                Case "vchUser35"
                    If Trim(Me.proUser35) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser35
                    End If
                Case "vchUser36"
                    If Trim(Me.proUser36) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser36
                    End If
                Case "vchUser37"
                    If Trim(Me.proUser37) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser37
                    End If
                Case "vchUser38"
                    If Trim(Me.proUser38) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser38
                    End If
                Case "vchUser39"
                    If Trim(Me.proUser39) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser39
                    End If
                Case "vchUser40"
                    If Trim(Me.proUser40) = "" Then
                        varComando.Parameters("@iValorId").Value = 0
                    Else
                        varComando.Parameters("@iValorId").Value = Me.proUser40
                    End If
            End Select
            
            Set varResultado = varComando.Execute
            
            If Not varResultado.EOF Then
                If IsNull(varResultado.Fields(0)) Then
                    MetValidarValor = True
                Else
                    parCampo = Me.proParametrosxProducto.Item(varContador).proEtiqueta
                    MetValidarValor = False
                    Exit Function
                End If
            Else
                MetValidarValor = True
            End If
        End If
    Next varContador
    
    Set varComando = Nothing
    Set varResultado = Nothing
    Exit Function
ErrManager:
    SubGMuestraError
End Function
