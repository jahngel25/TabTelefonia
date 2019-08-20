VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colReglasClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claReglasClas"
Attribute VB_Ext_KEY = "Member0" ,"claReglasClas"
Attribute VB_Ext_KEY = "Member1" ,"colClasificacion"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Modulo    : colReglasClasificacionMódulo de clase
' Fecha  : 27/09/2004 14:56
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : Atender la colección de la clase "claReglasClasificacion"
'---------------------------------------------------------------------------------------
Option Explicit

'Propiedad de conexión de la coleccion
Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

'Variables para consulta
Public proiClasificacionId As String
Public proiReglaId As String


Public Function FunGConsultaSeleccionados() As Boolean
Dim varResultado As ADODB.Recordset
Dim varReglasclasificacion As claReglasClasificacion
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultado = New ADODB.Recordset
        Set varReglasclasificacion = New claReglasClasificacion
        Set mCol = Nothing
        Set mCol = New Collection
        varResultado.Open "SELECT     iReglaId, vchDescripcionRegla " _
                                        & " From CT_REGLAS " _
                                        & " WHERE     (iReglaId  IN " _
                                        & " (SELECT     iReglaId " _
                                        & " From CT_ReglasClasificacion " _
                                       & " WHERE      iClasificacionID = " & Me.proiClasificacionId & "))", Me.proConexion
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
                varReglasclasificacion.ProDescripcionRegla = IIf(IsNull(varResultado("vchDescripcionRegla")), "", varResultado.Fields("vchDescripcionRegla"))
                varReglasclasificacion.proiReglaId = IIf(IsNull(varResultado("iReglaId")), "", varResultado.Fields("iReglaId"))
                Me.Add Me.proConexion, varReglasclasificacion.proiReglaId, _
                varReglasclasificacion.proiClasificacionId, varReglasclasificacion.ProDescripcionRegla
                
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        varResultado.Close
        FunGConsultaSeleccionados = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function FunGConsulta() As Boolean
Dim varResultado As ADODB.Recordset
Dim varReglasclasificacion As claReglasClasificacion
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultado = New ADODB.Recordset
        Set varReglasclasificacion = New claReglasClasificacion
        Set mCol = Nothing
        Set mCol = New Collection
        varResultado.Open "SELECT DISTINCT " & _
                                       "iClasificacionId, " & _
                                       "iReglaId " & _
                                       " FROM CT_ReglasClasificacion " & _
                                       " WHERE iClasificacionId =  " & Me.proiClasificacionId, Me.proConexion
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
                varReglasclasificacion.proiClasificacionId = IIf(IsNull(varResultado("iClasificacionId")), "", varResultado.Fields("iClasificacionId"))
                varReglasclasificacion.proiReglaId = IIf(IsNull(varResultado("iReglaId")), "", varResultado.Fields("iReglaId"))
               Me.Add Me.proConexion, varReglasclasificacion.proiReglaId, _
                varReglasclasificacion.proiClasificacionId, varReglasclasificacion.ProDescripcionRegla
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        
        FunGConsulta = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function FunGConsultaTodas() As Boolean
Dim varResultado As ADODB.Recordset
Dim varReglasclasificacion As claReglasClasificacion
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultado = New ADODB.Recordset
         Set varReglasclasificacion = New claReglasClasificacion
        Set mCol = Nothing
        Set mCol = New Collection
        varResultado.Open "SELECT     iReglaId, vchDescripcionRegla " _
                                        & " From CT_REGLAS " _
                                        & " WHERE     tiRecordStatus=1 ", Me.proConexion
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
                varReglasclasificacion.ProDescripcionRegla = IIf(IsNull(varResultado("vchDescripcionRegla")), "", varResultado.Fields("vchDescripcionRegla"))
                varReglasclasificacion.proiReglaId = IIf(IsNull(varResultado("iReglaId")), "", varResultado.Fields("iReglaId"))
                Me.Add Me.proConexion, varReglasclasificacion.proiReglaId, _
                varReglasclasificacion.proiClasificacionId, varReglasclasificacion.ProDescripcionRegla
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        
        FunGConsultaTodas = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function FunGConsultaFaltantes() As Boolean
Dim varResultado As ADODB.Recordset
Dim varReglasclasificacion As claReglasClasificacion
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultado = New ADODB.Recordset
         Set varReglasclasificacion = New claReglasClasificacion
        Set mCol = Nothing
        Set mCol = New Collection
        varResultado.Open "SELECT     iReglaId, vchDescripcionRegla " _
                                        & " From CT_REGLAS " _
                                        & " WHERE     (iReglaId NOT IN " _
                                        & " (SELECT     iReglaId " _
                                        & " From CT_ReglasClasificacion " _
                                       & " WHERE      iClasificacionID = " & Me.proiClasificacionId & "))", Me.proConexion _
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
                varReglasclasificacion.ProDescripcionRegla = IIf(IsNull(varResultado("vchDescripcionRegla")), "", varResultado.Fields("vchDescripcionRegla"))
                varReglasclasificacion.proiReglaId = IIf(IsNull(varResultado("iReglaId")), "", varResultado.Fields("iReglaId"))
                Me.Add Me.proConexion, varReglasclasificacion.proiReglaId, _
                varReglasclasificacion.proiClasificacionId, varReglasclasificacion.ProDescripcionRegla
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        
        FunGConsultaFaltantes = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function Add(proConexion As ADODB.Connection, proiReglaId As String, proiClasificacionId As String, Optional ProDescripcionRegla As String, Optional sKey As String) As claReglasClasificacion
    

    'create a new object
    Dim objNewMember As claReglasClasificacion
    Set objNewMember = New claReglasClasificacion


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proiReglaId = proiReglaId
    objNewMember.proiClasificacionId = proiClasificacionId
    objNewMember.proiReglaId = proiReglaId
    objNewMember.ProDescripcionRegla = ProDescripcionRegla

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claReglasClasificacion
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
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

