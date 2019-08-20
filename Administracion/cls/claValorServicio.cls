VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claValorServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proServicioSuplementarioId As String
Public proValorAnt As String
Public proValor As String
Public proDefault As String
Public protiRecordStatus As String

'Conexion
Public proConexion As ADODB.Connection

Public Function FunGInsertar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    varComando = "  INSERT INTO ct_valoresservicio  " & _
     "  (iServicioSuplementarioId, vchValor, bDefault, tiRecordStatus  ) " & _
     "  Values (" & Me.proServicioSuplementarioId & ",  '" & Me.proValor & "',  '" & proDefault & "'," & Me.protiRecordStatus & " )  "
    
    Me.proConexion.Execute varComando
    
    Set varResultado = New ADODB.Recordset

    
    Set varResultado = Nothing
    
    FunGInsertar = True
    
    Exit Function
ErrManager:
    FunGInsertar = False
    SubGMuestraError
End Function
Public Function FunGGuardar() As Boolean
    On Error GoTo ErrManager
    Dim strValorAux As String
    
    If Trim(Me.proValorAnt) = "" Then
     'Es un valor nuevo
        Me.protiRecordStatus = 1
        If Me.FunGInsertar Then
            FunGGuardar = True
        Else
            FunGGuardar = False
        End If
    Else
        If Trim(Me.proValor) = Trim(Me.proValorAnt) Then 'va a modificar el default que es lo único que se modifica
            Me.protiRecordStatus = 1
            If Me.FunGModificar Then
                FunGGuardar = True
            Else
                FunGGuardar = False
            End If
        Else 'quiere modificar la descripción para ello se debe eliminar ese valor y crear uno nuevo
            strValorAux = Trim(proValor)
            Me.protiRecordStatus = 0
            Me.proValor = Trim(Me.proValorAnt)
            If Me.FunGModificar Then
                FunGGuardar = True
            Else
                FunGGuardar = False
            End If
            Me.protiRecordStatus = 1
            Me.proValor = strValorAux
            If Me.FunGInsertar Then
                FunGGuardar = True
            Else
                FunGGuardar = False
            End If
        End If
    
    End If
    Exit Function
ErrManager:
    FunGGuardar = False
    SubGMuestraError
End Function
Public Function FunGModificar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "  Update  ct_valoresservicio " & _
     "  Set vchValor =  '" & Me.proValor & "', bDefault =  " & "'" & Me.proDefault & "'" & ", tiRecordStatus =  " & Me.protiRecordStatus & _
     "  Where vchValor =  '" & Me.proValor & "' AND iServicioSuplementarioId = " & proServicioSuplementarioId
    Me.proConexion.Execute varComando
    
    FunGModificar = True
    Exit Function
ErrManager:
    FunGModificar = False
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
    On Error GoTo ErrManager
    Dim strValorAux As String
    
    Me.protiRecordStatus = "0"
    If Me.FunGModificar Then
        FunGEliminar = True
    Else
        FunGEliminar = False
    End If
    
    Exit Function
ErrManager:
    FunGEliminar = False
    SubGMuestraError
End Function
