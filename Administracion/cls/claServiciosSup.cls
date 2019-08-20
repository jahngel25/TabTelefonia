VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claServiciosSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proiServicioSuplementarioId As String
Public provchNombreServicio As String
Public prochProductNumber As String
Public prochTipoServicio As String

'Conexion
Public proConexion As ADODB.Connection

Public Function FunGInsertar()
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    varComando = "  INSERT INTO ct_serviciossuplementarios  " & _
     "  ( vchNombreServicio, chProductNumber,chTipoServicio  ) " & _
     "  Values ( '" & Me.provchNombreServicio & "',  '" & Me.prochProductNumber & "','" & prochTipoServicio & "' )  "
    
    Me.proConexion.Execute varComando
    
    Set varResultado = New ADODB.Recordset

    varComando = "  Select Max(iServicioSuplementarioId) as ID   " & _
     "  From    ct_serviciossuplementarios  "
        
    varResultado.Open varComando, Me.proConexion
    Me.proiServicioSuplementarioId = IIf(IsNull(varResultado("ID")), "", varResultado.Fields("ID"))
    varResultado.Close
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
    
    varComando = "  Update  CT_ServiciosSuplementarios " & _
                 "  Set     vchNombreServicio =  '" & Me.provchNombreServicio & "', chProductNumber =  '" & Me.prochProductNumber & "'" & _
                 "  Where   iServicioSuplementarioId =  " & Me.proiServicioSuplementarioId
    
    Me.proConexion.Execute varComando
    
    FunGModificar = True
    Exit Function
ErrManager:
    FunGModificar = False
    SubGMuestraError
End Function

Public Function FunGGuardar() As Boolean
    On Error GoTo ErrManager
    
    If Trim(Me.proiServicioSuplementarioId) = "" Then
        If Me.FunGInsertar Then
            FunGGuardar = True
        Else
            FunGGuardar = False
        End If
    Else
        If Me.FunGModificar Then
            FunGGuardar = True
        Else
            FunGGuardar = False
        End If
    End If
    
    Exit Function
ErrManager:
    FunGGuardar = False
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "  Delete from  ct_serviciossuplementarios  " & _
     "  Where iServicioSuplementarioId =  " & Me.proiServicioSuplementarioId & "  AND chProductNumber =  " & "'" & Me.prochProductNumber & "'"
    Me.proConexion.Execute varComando
    
    FunGEliminar = True
    Exit Function
ErrManager:
    FunGEliminar = False
    SubGMuestraError
End Function

Function FunGConsulta() As Boolean
    On Error GoTo ErrManager
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    Set varResultado = New ADODB.Recordset
    
    varComando = "select iServicioSuplementarioId, vchNombreServicio,chProductNumber,chTipoServicio " & _
                 "from   CT_ServiciosSuplementarios " & _
                 "where  iServicioSuplementarioId = " & Me.proiServicioSuplementarioId
    
    varResultado.Open varComando, Me.proConexion
    
    If IsNull(varResultado.Fields(0)) Then
        Me.prochTipoServicio = "NINGUNO"
    Else
        Me.prochTipoServicio = varResultado.Fields(3)
    End If
    
    Set varResultado = Nothing
    FunGConsulta = True
    
    Exit Function
ErrManager:
    FunGConsulta = False
    SubGMuestraError
End Function
Private Sub Class_Initialize()
    On Error GoTo ErrorManager
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub
