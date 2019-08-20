VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claServiciosxReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proNovedadNumeroId As String
Public proServicioSuplementarioId As String
Public proTipoNovedadId As String

'Conexion
Public proConexion As ADODB.Connection

Public Function FunGInsertar()
Dim varComando As String
Dim varResultado As ADODB.Recordset
    
On Error GoTo ErrManager

  Set varResultado = New ADODB.Recordset
    
    'Debe verificar que no exista
    varResultado.Open "SELECT   Count(iNovedadNumeroId) numeroreg " & _
                  "FROM     ct_serviciosxreserva    " & _
                "WHERE    iNovedadNumeroId  = " & proNovedadNumeroId & " AND iServicioSuplementarioId =" & proServicioSuplementarioId, Me.proConexion

    If varResultado.EOF = False Then
        If CInt(varResultado("numeroreg")) = 0 Then
            'Si no existe se inserta
                    varComando = "  Insert into ct_serviciosxreserva  " & _
                            "  (iNovedadNumeroId, iServicioSuplementarioId, iTipoNovedadId) " & _
                            "  Values (" & proNovedadNumeroId & ",  " & proServicioSuplementarioId & ",  " & proTipoNovedadId & ") "
                proConexion.Execute varComando
        End If
    Else
        'debe eliminarlo y deja el último
        FunGEliminar
        'Si no existe se inserta
            varComando = "  Insert into ct_serviciosxreserva  " & _
                    "  (iNovedadNumeroId, iServicioSuplementarioId, iTipoNovedadId) " & _
                    "  Values (" & proNovedadNumeroId & ",  " & proServicioSuplementarioId & ",  " & proTipoNovedadId & ") "
        proConexion.Execute varComando
    End If
    Set varResultado = Nothing
    

FunGInsertar = True
Exit Function
ErrManager:
FunGInsertar = False
SubGMuestraError
End Function

Public Function FunGModificar() As Boolean
Dim varComando As String
    
On Error GoTo ErrManager
    
    varComando = "  Update  ct_serviciosxreserva " & _
                            "  Set iNovedadNumeroId = " & proNovedadNumeroId & ", " & _
                            " iServicioSuplementarioId = " & proServicioSuplementarioId & ", " & _
                            " iTipoNovedadId = " & proTipoNovedadId & _
                            " Where iNovedadNumeroId =  " & proNovedadNumeroId & _
                            " AND iServicioSuplementarioId =  " & proServicioSuplementarioId & _
                            " AND iTipoNovedadId =  " & proTipoNovedadId
                            proConexion.Execute varComando
    
FunGModificar = True
Exit Function
ErrManager:
FunGModificar = False
SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
Dim varComando As String
    
On Error GoTo ErrManager
    
    varComando = "  Delete from  ct_serviciosxreserva  " & _
                            " Where iNovedadNumeroId =  " & proNovedadNumeroId & _
                            " AND iServicioSuplementarioId =  " & proServicioSuplementarioId
    proConexion.Execute varComando
    
FunGEliminar = True
Exit Function
ErrManager:
FunGEliminar = False
SubGMuestraError
End Function


Private Sub Class_Initialize()
On Error GoTo ErrorManager
    
Exit Sub
ErrorManager:
SubGMuestraError
End Sub

