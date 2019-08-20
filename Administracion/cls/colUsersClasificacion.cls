VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colUsersClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proUserId As String
Public proUserName As String
Public proClasificacionId As String
Public proClasificacionDescripcion As String
Public proRecordStatus As String


Private varConexion As ADODB.Connection
Public proConexion As ADODB.Connection

Public proSeleccionados As Integer
Public proTodas As String

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proUserId As String, _
                    proUserName As String, _
                    proClasificacionId As String, proClasificacionDescripcion As String, _
                    Optional sKey As String) As claUsersClasificacion
    'create a new object
    Dim objNewMember As claUsersClasificacion
    Set objNewMember = New claUsersClasificacion

    
    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
        
    objNewMember.proClasificacionDescripcion = proClasificacionDescripcion
    objNewMember.proClasificacionId = proClasificacionId
    objNewMember.proUserId = proUserId
    objNewMember.proUserName = proUserName
    objNewMember.proRecordStatus = proRecordStatus
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claUsersClasificacion
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    Me.proSeleccionados = 0
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Private Function MetAbrirConexion() As Boolean
    On Error GoTo ErrManager
    
    Set varConexion = New ADODB.Connection
    varConexion.ConnectionString = Me.proConexion.ConnectionString
    varConexion.CommandTimeout = 0
    varConexion.ConnectionTimeout = 0
    
    varConexion.Open
    
    MetAbrirConexion = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
    


Public Function MetConsultarUserClasificacion() As Boolean
    Dim varComando As ADODB.Command
    Dim varComandoAux As String
    Dim varEntro As Boolean
    Dim varContador As Integer
    Dim varResultado As ADODB.Recordset
    Dim varUserCla As claUsersClasificacion
    Dim varCantidadClasificaciones As Integer
    Dim i As Integer
    On Error GoTo ErrManager
    
    'Limpia la coleccion
    If Me.Count > 0 Then
        For i = Me.Count To 1 Step -1
            Me.Remove (i)
        Next i
    End If
    
    Set varComando = New ADODB.Command
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CP_ConsultaUsersClasificacion"
    varComando.CommandTimeout = 0
    
    varComando.Parameters.Item("@iClasificacionId").Value = Me.proClasificacionId
    varComando.Parameters.Item("@chTodas").Value = Me.proTodas
    
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute
    
    Set varUserCla = New claUsersClasificacion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iClasificacionId")) Then
            varUserCla.proClasificacionId = ""
        Else
            varUserCla.proClasificacionId = Trim(varResultado.Fields("iClasificacionId"))
        End If
        
        If IsNull(varResultado.Fields("vchClasificacion")) Then
            varUserCla.proClasificacionDescripcion = ""
        Else
            varUserCla.proClasificacionDescripcion = Trim(varResultado.Fields("vchClasificacion"))
        End If
        
        If IsNull(varResultado.Fields("chUserId")) Then
            varUserCla.proUserId = ""
        Else
            varUserCla.proUserId = Trim(varResultado.Fields("chUserId"))
        End If
        
        If IsNull(varResultado.Fields("chUserName")) Then
            varUserCla.proUserName = ""
        Else
            varUserCla.proUserName = Trim(varResultado.Fields("chUserName"))
        End If
        
        Me.Add Me.proConexion, _
               varUserCla.proUserId, _
               varUserCla.proUserName, _
               varUserCla.proClasificacionId, _
               varUserCla.proClasificacionDescripcion
               
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varUserCla = Nothing
    
    MetConsultarUserClasificacion = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function FunGEliminarUserClasificacion(parItem As Integer) As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación lógica de un usuario por clasificacion
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Diana Buenhombre
' FECHA: 31 Enero 2006
'****************************************************************
On Error GoTo ErrorManager

        
        If MsgBox("Esta seguro de eliminar el usuario " & Me.Item(parItem).proUserName & " asignado a la clasificación " & Me.Item(parItem).proClasificacionDescripcion & "?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            FunGEliminarUserClasificacion = True
            Exit Function
        End If
    
        Set Me.Item(parItem).proConexion = Me.proConexion
        
               
        FunGEliminarUserClasificacion = Me.Item(parItem).FunGEliminar
        
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
