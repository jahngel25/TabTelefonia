VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claNovedadDetalleDatosProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proNovedadDetalleDatosProductoId As String
Public proTipoNovedadId As String
Public proDatosProductoId As String
Public proIncidentId As String
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
Public proContadorNumeros As Integer
Public proSeleccion As String
Public proConexion As ADODB.Connection

Public proParametrosxProducto As colParametroProducto

Public Function MetInsertar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    varComando = "Insert into CT_NovedadDetalleDatosProducto ( " & _
                 "          iTipoNovedadId, iDatosProductoId, iIncidentId, iDetalleDatosProductoId, " & _
                 "          chStatusId,     vchuser1, vchUser2, vchUser3, vchUser4, vchUser5,       " & _
                 "          vchUser6, vchUser7, vchUser8, vchUser9, vchUser10, vchUser11,           " & _
                 "          vchUser12, vchUser13,  vchUser14, vchUser15, vchUser16,       " & _
                 "          vchUser17, vchUser18, vchUser19, vchUser20, vchUser21,       " & _
                 "          vchUser22, vchUser23, vchUser24, vchUser25, vchUser26, vchUser27,       " & _
                 "          vchUser28, vchUser29, vchUser30, vchUser31, vchUser32, vchUser33,       " & _
                 "          vchUser34, vchUser35, vchUser36, vchUser37, vchUser38, vchUser39,       " & _
                 "          vchUser40, tiRecordStatus ) values (" & Me.proTipoNovedadId & ", " & _
                 Me.proDatosProductoId & ", " & Me.proIncidentId & ", " & Me.proDetalleDatosProductoId & ", '" & _
                 Me.proStatusId & "', '" & Me.proUser1 & "', '" & Me.proUser2 & "', '" & Me.proUser3 & "', '" & _
                 Me.proUser4 & "', '" & Me.proUser5 & "', '" & Me.proUser6 & "', '" & Me.proUser7 & "', '" & _
                 Me.proUser8 & "', '" & Me.proUser9 & "', '" & Me.proUser10 & "', '" & Me.proUser11 & "', '" & _
                 Me.proUser12 & "', '" & Me.proUser13 & "', '" & Me.proUser14 & "', '" & Me.proUser15 & "', '" & _
                 Me.proUser16 & "', '" & Me.proUser17 & "', '" & Me.proUser18 & "', '" & Me.proUser19 & "', '" & _
                 Me.proUser20 & "', '" & Me.proUser21 & "', '" & Me.proUser22 & "', '" & Me.proUser23 & "', '" & _
                 Me.proUser24 & "', '" & Me.proUser25 & "', '" & Me.proUser26 & "', '" & Me.proUser27 & "', '" & _
                 Me.proUser28 & "', '" & Me.proUser29 & "', '" & Me.proUser30 & "', '" & Me.proUser31 & "', '" & _
                 Me.proUser32 & "', '" & Me.proUser33 & "', '" & Me.proUser34 & "', '" & Me.proUser35 & "', '" & _
                 Me.proUser36 & "', '" & Me.proUser37 & "', '" & Me.proUser38 & "', '" & Me.proUser39 & "', '" & _
                 Me.proUser40 & "', " & Me.proRecordStatus & ") "
                 
    Me.proConexion.Execute varComando
    
    varComando = "Select Max(iNovedadDetalleDatosProductoId) from CT_NovedadDetalleDatosProducto " & _
                 "Where  iDatosProductoId = " & Me.proDatosProductoId & " and iIncidentId = " & Me.proIncidentId
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields(0)) Then
            Me.proNovedadDetalleDatosProductoId = 0
        Else
            Me.proNovedadDetalleDatosProductoId = Trim(varResultado.Fields(0))
        End If
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
    
    varComando = "Update    CT_NovedadDetalleDatosProducto     " & _
                 "Set   iTipoNovedadId = " & Me.proTipoNovedadId & ",  iDatosProductoId = " & Me.proDatosProductoId & ", " & _
                 "      iDetalleDatosProductoId = " & Me.proDetalleDatosProductoId & ", chStatusId = '" & Me.proStatusId & "', " & _
                 "      vchUser1 = '" & Me.proUser1 & "', vchUser2 = '" & Me.proUser2 & "', vchUser3 = '" & Me.proUser3 & "', " & _
                 "      vchUser4 = '" & Me.proUser4 & "', vchUser5 = '" & Me.proUser5 & "', vchUser6 = '" & Me.proUser6 & "', " & _
                 "      vchUser7 = '" & Me.proUser7 & "', vchUser8 = '" & Me.proUser8 & "', vchuser9 = '" & Me.proUser9 & "', " & _
                 "      vchUser10 = '" & Me.proUser10 & "', vchUser11 = '" & Me.proUser11 & "', vchUser12 = '" & Me.proUser12 & "', " & _
                 "      vchUser13 = '" & Me.proUser13 & "', vchUser14 = '" & Me.proUser14 & "', vchUser15 = '" & Me.proUser15 & "', " & _
                 "      vchUser16 = '" & Me.proUser16 & "', vchUser17 = '" & Me.proUser17 & "', vchUser18 = '" & Me.proUser18 & "', " & _
                 "      vchUser19 = '" & Me.proUser19 & "', vchUser20 = '" & Me.proUser20 & "', vchUser21 = '" & Me.proUser21 & "', " & _
                 "      vchUser22 = '" & Me.proUser22 & "', vchUser23 = '" & Me.proUser23 & "', vchUser24 = '" & Me.proUser24 & "', " & _
                 "      vchUser25 = '" & Me.proUser25 & "', vchUser26 = '" & Me.proUser26 & "', vchUser27 = '" & Me.proUser27 & "', " & _
                 "      vchUser28 = '" & Me.proUser28 & "', vchUser29 = '" & Me.proUser29 & "', vchUser30 = '" & Me.proUser30 & "', " & _
                 "      vchUser31 = '" & Me.proUser31 & "', vchUser32 = '" & Me.proUser32 & "', vchUser33 = '" & Me.proUser33 & "', " & _
                 "      vchUser34 = '" & Me.proUser34 & "', vchUser35 = '" & Me.proUser35 & "', vchUser36 = '" & Me.proUser36 & "', " & _
                 "      vchUser37 = '" & Me.proUser37 & "', vchUser38 = '" & Me.proUser38 & "', vchUser39 = '" & Me.proUser39 & "', " & _
                 "      vchUser40 = '" & Me.proUser40 & "', tiRecordStatus = " & Me.proRecordStatus & " " & _
                 "Where iNovedadDetalleDatosProductoId = " & Me.proNovedadDetalleDatosProductoId
    
    Me.proConexion.Execute varComando
    
    MetActualizar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "Delete from CT_NovedadDetalleDatosProducto " & _
                 "Where  iNovedadDetalleDatosProductoId = " & Me.proNovedadDetalleDatosProductoId
    
    Me.proConexion.Execute varComando
    
    MetEliminar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetGuardar() As Boolean
    On Error GoTo ErrManager
        
        If Val(Trim(Me.proNovedadDetalleDatosProductoId)) = 0 Then
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


Public Function MetValidarExistencia() As Boolean
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Public Sub FunGAumentarContador()
    proContadorNumeros = proContadorNumeros + 1
End Sub

Public Sub FunGDisminuirContador()
    proContadorNumeros = proContadorNumeros - 1
End Sub


